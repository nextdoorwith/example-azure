using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace GraphApiTestClient.Authentication
{
    // Microsoft Graph .NET Authentication Provider Library(Microsoft.Graph.Auth)で
    // 提供される認証プロバイダはpreview版なので採用見送り。
    // そのライブラリと同様にMSAL.NETライブラリを使用して認証プロバイダを実装する。
    public class MyAuthProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication _msalClient;

        private string[] _scopes;

        public MyAuthProvider(string clientId, string tenantId, string secret)
        {
            // Graph APIを使用する場合は固定
            _scopes = new string[] { "https://graph.microsoft.com/.default" };

            // Client Credentialsフローの場合は、機密クライアントアプリケーションを使用
            _msalClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(secret)
                .Build();
        }

        public async Task<string> GetAccessToken()
        {
            // TODO: 要件に応じてトークン取得のリトライ、キャッシングを実装
            var result = await _msalClient.AcquireTokenForClient(_scopes).ExecuteAsync();
            return result.AccessToken;
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("bearer", await GetAccessToken());
        }
    }
}