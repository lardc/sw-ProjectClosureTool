using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Web;
using System.Threading.Tasks;
using RestSharp;

namespace project_1
{

    public class Data {
        public String id { get; set; }
        public bool display { get; set; }
        public bool entities { get; set; }
        public String fields { get; set; }
        public bool member { get; set; }
    }

    class Program
    {
        private static object args;

        static void Main(string[] args)
        {
            TrelloRequestAsync("https://trello.com/b/I5Yr4tAu/software-management-automatization");
        }

        static async Task<string> TrelloRequestAsync(string url)
        {
            var tcs = new TaskCompletionSource<string>();
            try
            {
                var client = new RestClient(url);
                client.GetAsync(new RestRequest(), (response, handle) =>
                {
                    if ((int)response.StatusCode >= 400)
                    {
                        tcs.SetException(new Exception(response.StatusDescription));
                    }
                    else
                    {
                        tcs.SetResult(response.Content);
                    }
                });
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }

            //catch (InvalidOperationException e)
            //{

            //}
            //catch (HttpRequestException e)
            //{

            //}
            //catch (TaskCanceledException e)
            //{

            //}

            return await tcs.Task;
        }

        private static Task PostAsJsonAsync(string url, string id)
        {
            throw new NotImplementedException();
        }
    }
}