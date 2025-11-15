using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;

namespace H80
{
    internal class FlipMirror
    {
        private readonly HttpClient _httpClient = new HttpClient();

        public string Get(string url)
        {
            var requestUri = $"https://http://132.66.65.15/{Uri.EscapeDataString(url)}";
            var response = _httpClient.GetAsync(requestUri).Result;
            response.EnsureSuccessStatusCode();
            return response.Content.ReadAsStringAsync().Result;
        }

        public string CurrentCamera()
        {
            var response = Get("status");

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

        public void SelectCamera(string camera)
        {
            if (camera == CurrentCamera())
            {
                return;
            }

            string command = camera == "main" ? "set_port2" : "set_port1";
            Get(command);
        }
    }
}
