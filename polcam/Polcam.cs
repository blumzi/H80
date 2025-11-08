// Project: Class Library (.NET Framework 4.7.2), x86, "Register for COM interop"
// NuGet/Refs: ASCOM Platform (reference ASCOM.DriverAccess.dll, ASCOM.Utilities.dll), Microsoft.CSharp
// ProgId: h80polcam.Control
// After build (if not using VS "Register for COM interop"): regasm h80polcam.dll /codebase

using ASCOM.DriverAccess;     // from ASCOM Platform
using ASCOM.Utilities;        // for Chooser (optional)
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using ACP;    // for Plan/Target late-bound access
using static polcam.Globals;

namespace polcam
{
    [ComVisible(true)]
    [Guid("7E9B5C2E-7C28-4D3E-A5D5-5E7E2B58F0A1")]
    [ProgId("polcam")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Polcam
    {
        // ------------------- User Config (edit to match your rig) -------------------
        private const string Q550_PROGID = "ASCOM.QHYCCD_CAM2.Camera"; // ASCOM ProgID for QHY550P
        private const string FOCUSER_PROGID = "";                 // e.g. "ASCOM.Robofocus.Focuser"
        private const string FLIPMIRROR_PROGID = "";              // If flip mirror is an ASCOM Switch, put its ProgID here
        private const int FLIPMIRROR_PORT_INDEX_MAIN = 0;          // switch channel for main beam
        private const int FLIPMIRROR_PORT_INDEX_Q550 = 1;          // switch channel for Q550 branch

        private const bool USE_CHOOSERS_WHEN_EMPTY = true;         // pop Choosers if ProgID is empty
        private const string LOG_DIR = @"C:\\Wise\\Logs"; // ensure writeable

        // ------------------- Entry Point (called from ACP VBScript) -------------------
        // Signature kept as (object, object) to handle late-bound COM objects from ACP
        public void TargetEnd(object plan, object target)
        {
            logger.info("TargetEnd invoked.");

            // 1) Read polcam tag values from plan
            dynamic tag = FindTag(plan, "polcam", logger);
            double exposure = ReadDouble(tag, new[] { "exposure", "Exposure" }, 5.0, logger);
            int focuserOffset = (int)Math.Round(ReadDouble(tag, new[] { "focuser-offset", "FocuserOffset" }, 0, logger));
            short bin = (short)Math.Round(ReadDouble(tag, new[] { "bin" }, 1, logger));
            double gain = ReadDouble(tag, new[] { "gain" }, double.NaN, logger);
            double offset = ReadDouble(tag, new[] { "offset" }, double.NaN, logger);

            // 2) (Optional) get the primary image file path from target
            string primaryImagePath =
                FirstStringProp(target, new[] { "LastImagePath", "ImageFileName", "LastSavedFile", "FileName", "OutputFile" }, logger);

            if (!string.IsNullOrEmpty(primaryImagePath))
                logger.info($"Primary image path reported by target: {primaryImagePath}");
            else
                logger.info("Primary image path not exposed by target; continuing without it.");

            // 3) Flip mirror to Q550 branch (if configured)
            TryRouteFlipMirror(toQ550: true, log: logger);

            // 4) Nudge focuser by offset (if configured)
            if (!string.IsNullOrWhiteSpace(FOCUSER_PROGID) || USE_CHOOSERS_WHEN_EMPTY)
            {
                TryApplyFocuserOffset(focuserOffset, logger);
            }

            // 5) Capture with QHY550P
            string saveDir = EnsureDir(Path.Combine(GetDefaultAcpImagesDir(), "AuxCam"));
            string baseName = SafeBaseName(FirstStringProp(plan, new[] { "TargetName", "Name" }, logger),
                                            fallback: "Target");
            string fitsPath = UniquePath(saveDir, $"{baseName}_{UtcStamp()}_QHY550P_01.fits");

            CaptureQ550(exposure, bin, fitsPath, gain, offset, logger);

            // 6) Return flip mirror to main branch (optional)
            TryRouteFlipMirror(toQ550: false, log: logger);

            logger.info($"Done. Saved: {fitsPath}");
        }

        // ------------------- QHY550 capture via ASCOM -------------------
        private void CaptureQ550(double exposure, short bin, string fitsPath, double gain, double offset, Logger log)
        {
            string progId = Q550_PROGID;

            using (var cam = new Camera(progId))
            {
                log.info($"Connecting camera {progId}...");
                cam.Connected = true;

                try
                {
                    if (bin > 0) { cam.BinX = bin; cam.BinY = bin; }

                    // If driver exposes gain/offset as doubles/ints via Camera properties
                    TrySetProperty(cam, "Gain", gain, log);
                    TrySetProperty(cam, "Offset", offset, log);

                    log.info($"StartExposure {exposure}s, bin={bin}");
                    cam.StartExposure(exposure, true);
                    while (!cam.ImageReady) Thread.Sleep(150);

                    var img = cam.ImageArray as Array; // 2D [y,x]
                    if (img == null)
                        throw new InvalidOperationException("Camera.ImageArray returned null or unexpected type.");

                    // Dimensions
                    int ny = img.GetLength(0);
                    int nx = img.GetLength(1);

                    WriteFits16Mono(fitsPath, img, nx, ny, header:
                        h => {
                            h.CardS("OBJECT", ""); // we’ll fill from file name
                            h.CardS("INSTRUME", "QHY550P");
                            h.CardS("DATE-OBS", UtcStamp());
                            h.CardS("EXPTIME", exposure.ToString("0.###"));
                            //h.CardK("XBINNING", bin.ToString());
                            if (!double.IsNaN(gain)) h.CardS("EGAIN", gain.ToString());
                            if (!double.IsNaN(offset)) h.CardS("EOFFSET", offset.ToString());
                        });
                }
                finally
                {
                    cam.Connected = false;
                }
            }
        }

        // ------------------- Focuser offset -------------------
        private void TryApplyFocuserOffset(int offset, Logger log)
        {
            if (offset == 0) { log.info("Focuser offset = 0 (skipped)."); return; }

            string progId = FOCUSER_PROGID;
            if (string.IsNullOrWhiteSpace(progId) && USE_CHOOSERS_WHEN_EMPTY)
            {
                var chooser = new Chooser { DeviceType = "Focuser" };
                progId = chooser.Choose(null);
            }
            if (string.IsNullOrWhiteSpace(progId)) { log.info("No focuser configured."); return; }

            using (var foc = new Focuser(progId))
            {
                foc.Connected = true;
                try
                {
                    int start = foc.Position;
                    int dest = Math.Max(0, start + offset);
                    log.info($"Focuser move: {start} -> {dest}");
                    foc.Move(dest);
                }
                finally { foc.Connected = false; }
            }
        }

        // ------------------- Flip mirror routing (as ASCOM Switch) -------------------
        private void TryRouteFlipMirror(bool toQ550, Logger log)
        {
            string progId = FLIPMIRROR_PROGID;
            if (string.IsNullOrWhiteSpace(progId)) { log.info("No flip mirror configured."); return; }

            using (var sw = new ASCOM.DriverAccess.Switch(progId))
            {
                sw.Connected = true;
                try
                {
                    int index = toQ550 ? FLIPMIRROR_PORT_INDEX_Q550 : FLIPMIRROR_PORT_INDEX_MAIN;
                    int count = sw.MaxSwitch;
                    if (index < 0 || index >= count) throw new IndexOutOfRangeException($"Flip mirror index {index} out of 0..{count - 1}");
                    log.info($"Flip mirror -> channel {index}");
                    // Some devices are momentary; most beam selectors are latched booleans per channel
                    for (short i = 0; i < count; i++)
                        SafeSetSwitch(sw, i, i == index);
                }
                finally { sw.Connected = false; }
            }
        }

        private void SafeSetSwitch(ASCOM.DriverAccess.Switch sw, short index, bool state)
        {
            try { if (sw.CanWrite(index)) sw.SetSwitch(index, state); }
            catch { /* ignore */ }
        }

        // ------------------- Plan tag parsing (late-bound COM) -------------------
        private dynamic FindTag(object plan, string wantedName, Logger log)
        {
            if (plan == null) return null;
            var tags = TryGet(plan, "Tags");
            if (tags == null) { log.info("Plan has no Tags collection."); return null; }

            foreach (var t in (System.Collections.IEnumerable)tags)
            {
                var name = (TryGet(t, "Name") ?? "").ToString();
                if (string.Equals(name, wantedName, StringComparison.OrdinalIgnoreCase))
                    return t;
            }
            log.info($"Tag '{wantedName}' not found.");
            return null;
        }

        private double ReadDouble(object tag, string[] keys, double def, Logger log)
        {
            if (tag == null) return def;
            foreach (var k in keys)
            {
                var val = TagValue(tag, k);
                if (val != null && double.TryParse(val.ToString(), out double d)) return d;
            }
            return def;
        }

        private object TagValue(object tag, string key)
        {
            // Try common access patterns seen in ACP tags
            return TryCall(tag, "Item", key) ??
                   TryGet(tag, key) ??
                   (TryGet(TryGet(tag, "Values"), key)) ??
                   TryCall(tag, "Get", key);
        }

        private string FirstStringProp(object obj, string[] names, Logger log)
        {
            foreach (var n in names)
            {
                var v = TryGet(obj, n);
                if (v != null)
                {
                    var s = v.ToString();
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                }
            }
            return string.Empty;
        }

        // ------------------- Late-bound helpers -------------------
        private object TryGet(object o, string prop)
        {
            try { return o?.GetType().InvokeMember(prop, System.Reflection.BindingFlags.GetProperty, null, o, null); }
            catch { try { return o?.GetType().InvokeMember(prop, System.Reflection.BindingFlags.GetField, null, o, null); } catch { return null; } }
        }

        private object TryCall(object o, string method, params object[] args)
        {
            try { return o?.GetType().InvokeMember(method, System.Reflection.BindingFlags.InvokeMethod, null, o, args); }
            catch { return null; }
        }

        private void TrySetProperty(object o, string prop, double value, Logger log)
        {
            if (double.IsNaN(value)) return;
            try { o?.GetType().InvokeMember(prop, System.Reflection.BindingFlags.SetProperty, null, o, new object[] { value }); log.info($"Set {prop}={value}"); }
            catch { log.info($"Driver does not expose {prop} (skipped)."); }
        }

        // ------------------- FITS writer (16-bit mono, big-endian) -------------------
        private void WriteFits16Mono(string path, Array img, int nx, int ny, Action<FitsHeader> header)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read))
            using (var bw = new BinaryWriter(fs, Encoding.ASCII))
            {
                var h = new FitsHeader();
                h.CardK("SIMPLE", "T", "Standard FITS");
                h.CardK("BITPIX", "16", "16-bit signed ints");
                h.CardK("NAXIS", "2", null);
                h.CardK("NAXIS1", nx.ToString(), null);
                h.CardK("NAXIS2", ny.ToString(), null);
                h.CardK("BZERO", "32768", "Unsigned offset");
                h.CardK("BSCALE", "1", null);
                header?.Invoke(h);
                h.CardEnd();

                var hdr = h.GetBytes();
                bw.Write(hdr);
                // pad to 2880 bytes
                int pad = (2880 - (hdr.Length % 2880)) % 2880;
                if (pad > 0) bw.Write(new byte[pad]);

                // pixels: big-endian signed 16-bit with BZERO=32768
                for (int y = 0; y < ny; y++)
                {
                    for (int x = 0; x < nx; x++)
                    {
                        int v = Convert.ToInt32(img.GetValue(y, x) ?? 0);
                        if (v < 0) v = 0; if (v > 65535) v = 65535;
                        int s = v - 32768; if (s < 0) s += 65536; // two's complement
                        byte hi = (byte)((s >> 8) & 0xFF);
                        byte lo = (byte)(s & 0xFF);
                        bw.Write(hi); bw.Write(lo);
                    }
                }
            }
        }

        // ------------------- Small utils -------------------
        private static string GetDefaultAcpImagesDir()
        {
            string pub = Environment.GetEnvironmentVariable("PUBLIC");
            if (string.IsNullOrWhiteSpace(pub)) pub = @"C:\\Users\\Public";
            return Path.Combine(pub, "Documents", "ACP Images");
        }

        private static string EnsureDir(string p) { Directory.CreateDirectory(p); return p; }
        private static string SafeBaseName(string s, string fallback)
        {
            if (string.IsNullOrWhiteSpace(s)) s = fallback;
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            return s.Replace(':', '_').Replace('/', '_').Replace('\\', '_');
        }
        private static string DateStamp() => DateTime.Now.ToString("yyyy-MM-dd");
        private static string UtcStamp() => DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'");
        private static string UniquePath(string dir, string name)
        {
            string p = Path.Combine(dir, name);
            if (!File.Exists(p)) return p;
            string stem = Path.GetFileNameWithoutExtension(name);
            string ext = Path.GetExtension(name);
            for (int i = 2; i < 100; i++)
            {
                string q = Path.Combine(dir, $"{stem}_{i:D2}{ext}");
                if (!File.Exists(q)) return q;
            }
            return p;
        }
    }

    // ------------------- FITS header builder -------------------
    internal class FitsHeader
    {
        private readonly MemoryStream _ms = new MemoryStream();
        public void CardK(string key, string val, string comment)
        {
            string s = key.PadRight(8) + "= " + val.PadRight(20);
            if (!string.IsNullOrEmpty(comment)) s = s.PadRight(80 - (comment.Length + 2)) + "/ " + comment;
            Write80(s);
        }
        public void CardS(string key, string val)
        {
            string s = key.PadRight(8) + "= '" + val.Replace("'", "''") + "'";
            Write80(s);
        }
        public void CardEnd() => Write80("END");
        private void Write80(string s)
        {
            var b = Encoding.ASCII.GetBytes((s.Length > 80 ? s.Substring(0, 80) : s.PadRight(80)));
            _ms.Write(b, 0, b.Length);
        }
        public byte[] GetBytes() => _ms.ToArray();
    }
}
