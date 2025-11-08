using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Threading.Tasks;

namespace polcam
{
    public class Flip
    {
        private readonly string url_base = "http://132.66.65.15/";

        public Flip() {
        }

        public async Task<string> Get(string what)
        {
            using (var client = new HttpClient())
            {
                var getResponse = await client.GetAsync($"{this.url_base}/{what}");
                string content = await getResponse.Content.ReadAsStringAsync();
                return content.Trim();
            }
        }

        public async Task<string> Post(string what)
        {
            using (var client = new HttpClient())
            {
                var getResponse = await client.PostAsync($"{this.url_base}/{what}", null);
                string content = await getResponse.Content.ReadAsStringAsync();
                return content.Trim();
            }
        }

        public string Status()
        {
            return Get("status").Result;
        }

        public void SelectPolarizer()
        {
            var task = Post("set_port1");
            task.Wait();
        }

        public void SelectMainCamera()
        {
            var task = Post("set_port2");
            task.Wait();
        }
    };

}
