using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace GoogleCalender_Demo
{

    public class Program
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        
        static void Main(string[] args)
        {
            //專案名稱
            string ApplicationName = "Google Calendar API .NET Demo";

            //取得Cred Json路徑
            var FilePath = GetCredentialsPath();

            //取得權限
            var Scopes = GetScopes();

            //OAuth
            UserCredential Credential;
            using (var Stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                //產生通過驗證的Token.Josn
                string CredPath = "token.json";
                Credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(Stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(CredPath, true)
                    ).Result;
            }

            var @event = BookingCalendar();

            #region 修改Calendar Id
            //Calendar類別,僅Booking個人行事曆 CalendarId = Private
            string CalendarId = "....";
            #endregion

            //初始化設定
            var Service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = ApplicationName,
            });

            //Insert 欲Booking的資訊(即BookingCalendar()的設定值)及Calendar分類
            //Events也包含其餘Delete、Get、Update功能
            var query = Service.Events.Insert(@event, CalendarId).Execute();

            //Booking完成後會回傳Link 點開後即可看見Booking資訊
            Console.WriteLine(query.HtmlLink);
            Console.Read();
        }

        private static Event BookingCalendar()
        {
            //設定Time Zone
            string TimeZone = @"Asia/Taipei";
            Event @event = new Event();

            //Calendar Title & Desc
            @event.Summary = "Test";
            @event.Description = "Hauser Google Calendar Api Test";

            //起訖時間設定
            EventDateTime StartEventDateTime = new EventDateTime();
            StartEventDateTime.TimeZone = TimeZone;
            StartEventDateTime.DateTime = DateTime.Now;

           
            EventDateTime EndEventDateTime = new EventDateTime();
            EndEventDateTime.TimeZone = TimeZone;
            EndEventDateTime.DateTime = DateTime.Now.AddHours(1);

            @event.Start = StartEventDateTime;
            @event.End = EndEventDateTime;

            #region 與會人邀請
            List<string> AttendeesEmail = new List<string>();
            var Attendees = new List<EventAttendee>();
            foreach (var item in AttendeesEmail)
            {
                Attendees.Add(new EventAttendee()
                {
                    Email = item
                });
            }

            if (Attendees.Count > 0)
            {
                @event.Attendees = Attendees;
            }
            #endregion



            return @event;
            

        }


        /// <summary>
        /// 取得Json路徑
        /// </summary>
        /// <returns></returns>
        private static string GetCredentialsPath()
        {
            string CredPath =  "./Credentials/" + "credentials.json";
            return CredPath;
        }

        /// <summary>
        /// 設定Calendar
        /// 預設最大值
        /// </summary>
        /// <returns></returns>
        private static string[] GetScopes()
        {
            //可調整Calendar權限
            string[] Scopes =
               {
                    CalendarService.Scope.Calendar
                };
            return Scopes;
        }
    }
}

#region Step
/* 1.
 * 到下列網址Enable Calendar API
 * 並下載Json檔案放置到Credentials資料夾內
 * https://developers.google.com/calendar/quickstart/dotnet 
 * 
 * 2.
 * Install-Package Google.Apis.Calendar.v3
 * 
 * 3.
 * 依照個人情況修改BookingCalendar()及Region區塊
 *
 * 4.
 * 執行程式並利用Return的Link來檢查是否Booking成功
*/
#endregion