using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace Sowfin.API.Lib.ErrorHandling
{
    public class ErrorResponse
    {
        private HttpStatusCode code;
        private string message;
        public ErrorResponse(HttpStatusCode code, string message)
        {
            this.message = message;
            this.code = code;
        }

        public Task<HttpResponseMessage> ReturnError(HttpStatusCode code, string message)
        {
            var response = new HttpResponseMessage(code);
            response.Content = new StringContent(message);
            return Task.FromResult(response);

        }
    }
}
