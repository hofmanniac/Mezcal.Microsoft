using Mezcal.Commands;
using SP = Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Microsoft.SharePoint.Client;
using System.Security;
using Mezcal.Connections;

namespace Mezcal.Microsoft.Office
{
    public class SharePointDownload : ICommand
    {
        public void Process(JObject command, Context context)
        {
            var siteurl = JSONUtil.GetText(command, "url");
            var clientContext = new SP.ClientContext(siteurl);
            string serverRelativeUrlOfFile = JSONUtil.GetText(command, "sourcefile");
            string fileDestinationPath = JSONUtil.GetText(command, "destinationfile");

            clientContext.Credentials = CredentialCache.DefaultNetworkCredentials;
            var un = JSONUtil.GetText(command, "un");
            var p = new SecureString();
            var pw = JSONUtil.GetText(command, "pw");
            foreach(var c in pw) { p.AppendChar(c); }

            // doesn't appear to work for non onmicrosoft.com (corp) credentials
            // need cookies?
            clientContext.Credentials = new SharePointOnlineCredentials(un, p);
            
            using (SP.FileInformation sharePointFile =
                SP.File.OpenBinaryDirect(clientContext, serverRelativeUrlOfFile))
            {
                using (Stream destFile = System.IO.File.OpenWrite(fileDestinationPath))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int byteReadInLastRead;
                    while ((byteReadInLastRead = sharePointFile.Stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        destFile.Write(buffer, 0, byteReadInLastRead);
                    }
                }
            }

            System.Diagnostics.Process.Start(fileDestinationPath);
        }

        private CookieContainer GetAuthCookies(Uri webUri, string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (var c in password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(userName, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(webUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(webUri, authCookie);
            return cookieContainer;
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
