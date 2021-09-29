using Microsoft.AspNet.SignalR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace hospital.Models
{
    //https://blog.darkthread.net/blog/upload-progress-bar-w-signalr/
    //用於進度條顯示
    public class UploaderHub : Hub
    {
        //從其他類別呼叫需用以下方法取得UploaderHub Instance
        static IHubContext HubContext =
                        GlobalHost.ConnectionManager.GetHubContext<UploaderHub>();

        public static void UpdateProgress(string connId, string name, float percentage,
                                          string progress, string message = null)
        {
            HubContext.Clients.Client(connId)
                .updateProgress(name, percentage, progress, message);
        }
    }
}