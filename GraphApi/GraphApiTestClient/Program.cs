using GraphApiTestClient.Authentication;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GraphApiTestClient
{
    class Program
    {
        // Azure AD B2Cに登録したアプリのクライアントID
        const string ClientId = "11111111-2222-3333-4444-555555555555";
        // Azure AD B2CテナントのID
        const string TenantId = "mytenant.onmicrosoft.com";
        // 本来はKey Vault等の安全なストアを推奨
        const string Secret = "xxxxxxxxxxxxxxxxxxx";
        // カスタム属性操作のための"b2c-extensions-app"アプリのクライアントID
        const string B2CExtClientId = "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz";
        // 拡張属性に含まれるカスタム属性名の定義
        // 形式: "extension_{guid}_{属性名}"
        static string _exCusAttrPrefix = $"extension_{B2CExtClientId.Replace("-", "")}_";
        static string _exCusAttrCustomString = _exCusAttrPrefix + "customString";
        static string _exCusAttrCustomInt = _exCusAttrPrefix + "customInt";
        static string _exCusAttrCustomBoolean = _exCusAttrPrefix + "customBoolean";

        static async Task Main(string[] args)
        {
            // クライアントの認証に使用する認証プロバイダを生成
            var authProvider = new MyAuthProvider(ClientId, TenantId, Secret);

            // Microsoft Graphを操作するためのクライアントの生成
            var graphClient = new GraphServiceClient(authProvider);

            // ユーザアカウント一覧の取得
            await GetUserList(graphClient);

            // ユーザアカウントの登録
            var id = await CreateUser(graphClient);
            await GetUser(graphClient, id);

            // ユーザアカウントの更新
            await UpdateUser(graphClient, id);
            await GetUser(graphClient, id);

            // ユーザアカウント(パスワード)の更新
            await UpdateUserPassword(graphClient, id, "newpassword!");

            // ユーザアカウントの削除
            await DeleteUser(graphClient, id);
        }

        static async Task GetUserList(GraphServiceClient client)
        {
            Console.WriteLine("getting user list...");

            var resultList = new List<User>();

            // 最初のページ分のユーザ一覧を取得
            var usersPage = await client.Users
                .Request()
                .Select(e => new
                {
                    e.Id,
                    e.DisplayName
                })
                .OrderBy("displayName")
                .GetAsync();
            resultList.AddRange(usersPage.CurrentPage);

            // 次のページ以降のユーザ一覧を取得
            while(usersPage.NextPageRequest != null)
            {
                usersPage = await usersPage.NextPageRequest.GetAsync();
                resultList.AddRange(usersPage.CurrentPage);
            }

            Console.WriteLine($"User Count: {resultList.Count}");
            foreach (var u in resultList)
            {
                Console.WriteLine($"ObjectId={u.Id}, displayName={u.DisplayName}");
            }
            Console.WriteLine();
        }

        static async Task GetUser(GraphServiceClient client, string id)
        {
            Console.WriteLine("getting user...");

            // カスタム属性に値を指定する場合はUser.AdditionalDataを使用するが
            // 値を取得する場合は、User.AdditionalDataでは取得不可
            // 回避策として、次の2つの方法が考えられるが、
            // わざわざ要求を再送信するのも手間なので(1)の方法を使用する。
            // (1) 文字列でカスタム属性の名称を指定(結果として他の属性目も文字列指定)
            // (2) User.Extensionsを別途取得する。
            //     (client.Users[id].Extensions.Request().GetAsync())

            var result = await client.Users[id]
                .Request()
                .Select(
                    nameof(User.Id)
                    + "," + nameof(User.DisplayName)
                    + "," + nameof(User.Identities)
                    + "," + nameof(User.UserPrincipalName)
                    + "," + nameof(User.MailNickname)
                    + "," + nameof(User.AccountEnabled)
                    + "," + nameof(User.PasswordProfile)
                    + "," + nameof(User.PasswordPolicies)
                    + "," + nameof(User.OtherMails)
                    + "," + nameof(User.EmployeeId)
                    + "," + _exCusAttrCustomString
                    + "," + _exCusAttrCustomInt
                    + "," + _exCusAttrCustomBoolean
                )
                .GetAsync();
                
            ShowUser(result);
        }

        static async Task<string> CreateUser(GraphServiceClient client)
        {
            Console.WriteLine("creating user...");

            // 基本情報
            var newGuid = Guid.NewGuid().ToString();
            var user = new User()
            {
                DisplayName = "テストユーザ１",
                // サインインユーザ名、サインインメールアドレス
                // (ポータルだと1つ以上必須）
                Identities = new List<ObjectIdentity>()
                {
                    new ObjectIdentity()
                    {
                        SignInType = "userName",
                        Issuer = TenantId,
                        IssuerAssignedId = "testuser01"
                    },
                    new ObjectIdentity()
                    {
                        SignInType = "emailAddress",
                        Issuer = TenantId,
                        IssuerAssignedId = "testuser01@example.com"
                    }
                },
                // ポータルと同様にGUIDを設定
                UserPrincipalName = $"{newGuid}@{TenantId}",
                // ポータルと同様にGUIDを設定
                MailNickname = newGuid,
                // ポータルの[サインインのブロック]に対応(意味が逆)
                AccountEnabled = true,
                // ポータルと同様に、次回サインイン時のパスワード変更を要求しない
                PasswordProfile = new PasswordProfile()
                {
                    Password = "password",
                    ForceChangePasswordNextSignIn = false
                },
                // ポータル同様に、無期限、複雑性要求なしを指定
                PasswordPolicies = 
                    "DisablePasswordExpiration, " + 
                    "DisableStrongPassword",
            };

            // 拡張属性に含まれる既定の属性・カスタム属性の指定サンプル

            // 拡張属性に含まれる従業員IDはプロパティ指定可
            user.EmployeeId = "A12346";

            // 拡張属性に含まれるカスタム属性はAdditionalDataで指定
            user.AdditionalData = new Dictionary<string, object>()
            {
                //["employeeId"] = "A12346",
                [_exCusAttrCustomString] = "ハロー",
                [_exCusAttrCustomInt] = 123456,
                [_exCusAttrCustomBoolean] = true
            };

            // ユーザアカウントの作成
            var result = await client.Users
                .Request()
                .AddAsync(user);

            Console.WriteLine($"created: Id={result.Id}");
            Console.WriteLine();

            return result.Id;
        }

        static async Task<string> CreateSimpleUser(GraphServiceClient client)
        {
            Console.WriteLine("creating simple user...");

            // 基本情報
            var newGuid = Guid.NewGuid().ToString();
            var user = new User()
            {
                DisplayName = "テストユーザ２",
                UserPrincipalName = $"{newGuid}@{TenantId}",
                MailNickname = newGuid,
                AccountEnabled = true,
                PasswordProfile = new PasswordProfile()
                {
                    Password = "P@ssword!",
                }
            };

            // ユーザアカウントの作成
            var result = await client.Users
                .Request()
                .AddAsync(user);
            Console.WriteLine($"created: Id={result.Id}");
            Console.WriteLine();

            return result.Id;
        }

        static async Task UpdateUser(GraphServiceClient client, string id)
        {
            Console.WriteLine("updating user...");

            // 基本情報
            var user = new User()
            {
                DisplayName = "テストユーザ１更新",
                // ユーザ特定のキーとなるサインインユーザ名とメールアドレスを変更
                Identities = new List<ObjectIdentity>()
                {
                    new ObjectIdentity()
                    {
                        SignInType = "userName",
                        Issuer = TenantId,
                        IssuerAssignedId = "testuser01update"
                    },
                    new ObjectIdentity()
                    {
                        SignInType = "emailAddress",
                        Issuer = TenantId,
                        IssuerAssignedId = "testuser01update@example.com"
                    }
                },
                // ユーザ特定のキーとなるUPNを変更
                UserPrincipalName = $"test1234567890@{TenantId}",
                // サインインのブロック
                AccountEnabled = false,
                // 連絡用メールアドレスの変更サンプル
                OtherMails = new String[]{ 
                    "contact1@example.com", "contact2@example.com"
                }
            };

            // 拡張属性のカスタム属性の変更サンプル
            user.AdditionalData = new Dictionary<string, object>()
            {
                [_exCusAttrCustomString] = "ハローupdate",
                [_exCusAttrCustomInt] = 654321,
                [_exCusAttrCustomBoolean] = false
            };

            // ユーザアカウントの更新
            var result = await client.Users[id]
                .Request()
                .UpdateAsync(user);

            Console.WriteLine($"updated: Id={id}");
            Console.WriteLine();
        }

        static async Task UpdateUserPassword(
            GraphServiceClient client, string id, string password)
        {
            Console.WriteLine("updating user password...");

            var user = new User()
            {
                PasswordProfile = new PasswordProfile()
                {
                    Password = password,
                    ForceChangePasswordNextSignIn = false
                }
            };
            await client.Users[id]
                .Request()
                .UpdateAsync(user);

            Console.WriteLine($"updated: Id={id}");
            Console.WriteLine();
        }

        static async Task DeleteUser(GraphServiceClient client, string id)
        {
            Console.WriteLine("deleting user...");
            await client.Users[id]
                .Request()
                .DeleteAsync();
            Console.WriteLine($"deleted: Id={id}");
            Console.WriteLine();
        }

        static void ShowUser(User user)
        {
            const string FMT = "> {0,-30}: {1}";

            Console.WriteLine(FMT, "Id" , user.Id);
            Console.WriteLine(FMT, "DisplayName", user.DisplayName);
            foreach(var identity in user.Identities)
            {
                Console.WriteLine(FMT,
                    "Identities." + identity.SignInType, identity.IssuerAssignedId);
            }
            Console.WriteLine(FMT, "UserPrincipalName", user.UserPrincipalName);
            Console.WriteLine(FMT, "MailNickname", user.MailNickname);
            Console.WriteLine(FMT, "AccountEnabled", user.AccountEnabled);
            if( user.PasswordProfile != null)
            {
                Console.WriteLine(FMT, "PasswordProfile.Force...NextSignIn", 
                    user.PasswordProfile.ForceChangePasswordNextSignIn);
                Console.WriteLine(FMT, "PasswordProfile.Force...NextSignInWithMfa", 
                    user.PasswordProfile.ForceChangePasswordNextSignInWithMfa);
            }
            Console.WriteLine(FMT, "PasswordPolicies", user.PasswordPolicies);
            if( user.OtherMails != null)
            {
                int i = 0;
                foreach(var mail in user.OtherMails)
                {
                    Console.WriteLine(FMT, $"OtherMails[{i}]", mail);
                    i++;
                }
            }
            Console.WriteLine(FMT, "EmployeeId", user.EmployeeId);
            foreach (var pair in user.AdditionalData)
            {
                if (!pair.Key.StartsWith(_exCusAttrPrefix)) continue; // 関係ないものは除外
                Console.WriteLine(FMT, $"AdditionalData.{pair.Key}", pair.Value);
            }

            Console.WriteLine();
        }

    }
}
