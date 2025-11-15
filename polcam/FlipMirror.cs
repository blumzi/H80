using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace H80
{
    internal class FlipMirror
    {
        private readonly HttpClient _httpClient = new HttpClient();

        public async Task<string> Get(string url)
        {
            var requestUri = $"https://http://132.66.65.15/{Uri.EscapeDataString(url)}";
            var response = await _httpClient.GetAsync(requestUri);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        public async Task<string> CurrentCamera()
        {
            var response = await Get("status");

            if (response.Contains("S_AIN_BOUT"))
            {
                return "main";
            }
            else if (response.Contains("S_AOUT_BOUT"))
            {
                return "polar";
            }
            else
            {
                return "unknown";
            }
        }

        public async Task SelectCamera(string camera)
        {
            if (camera == CurrentCamera().Result)
            {
                return;
            }

            string command = camera == "main" ? "set_port2" : "set_port1";
            await Get(command);
        }
    }
}
