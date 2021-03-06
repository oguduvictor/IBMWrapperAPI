﻿using IBM.WatsonDeveloperCloud.PersonalityInsights.v3;
using IBM.WatsonDeveloperCloud.PersonalityInsights.v3.Model;
using IBMWrapperAPI.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Tweetinvi;
using Tweetinvi.Models;

namespace IBMWrapperAPI.Controllers
{
    /// <summary>
    /// Wrapper Controller
    /// </summary>
    [Route("api/[controller]/[action]")]
    [EnableCors("MyPolicy")]
    public class PersonalityController : Controller
    {
        private readonly string IBMId;
        private readonly string IBMPassword;
        private readonly PersonalityInsightsService _personalityInsights;
        private const string CONSUMER_KEY = "W4xk2V4Yc4hgHrJdDHhnwahQ1";
        private const string CONSUMER_SECRET = "xZU1ThN5wjFzC9pYj5NxBnjk6k6ohssGpTLa4qp80eA0ruDPsW";
        private const string ACCESS_TOKEN = "1289180162-7pFDG8ItGfyExFZdEOZj2uKuyKJ83BaJaUrkiWq";
        private const string ACCESS_TOKEN_SECRET = "UOLjtrHylaSA9IQ2JFl1T9efJEWXJ219JcmYBBvzmcJXv";
        private readonly TwitterCredentials twitterCredentials;
        /// <summary>
        /// Wrapper Controller Constructor
        /// </summary>
        public PersonalityController()
        {
            IBMId = "f7de21f9-1eee-4e47-9ff2-1040d6c762cb";
            IBMPassword = "WJZs7TfsDzdI";
            _personalityInsights = new PersonalityInsightsService(IBMId, IBMPassword, "2017-10-13");
            twitterCredentials = new TwitterCredentials(CONSUMER_KEY, CONSUMER_SECRET, ACCESS_TOKEN, ACCESS_TOKEN_SECRET);
            //Auth.SetUserCredentials(CONSUMER_KEY, CONSUMER_SECRET, ACCESS_TOKEN, ACCESS_TOKEN_SECRET);
            //Auth.SetCredentials(twitterCredentials);
        }

        /// <summary>
        /// Method for analyzing personality characteristics based on how a person writes.
        /// </summary>
        /// <param name="data">Text data to analyze</param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult AnalyseText([FromBody] string data)
        {
            try
            {
                if (data == null)
                {
                    throw new ArgumentException("Enter Valid String of atleast 100 characters to analyse");
                }

                return AnalyseData(data);
            }
            catch (Exception ex)
            {
                return StatusCode(400, ex.Message);
            }
        }
        
        /// <summary>
        /// Method to Extract personality characteristics based on how a person writes from a file.
        /// </summary>
        /// <param name="file">file to analyze</param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult AnalyseFile(IFormFile file)
        {
            try
            {
                if (file == null)
                {
                    throw new ArgumentException("File is required");
                }

                var data = ReadFileContent(file);

                return AnalyseData(data);
            }
            catch (Exception ex)
            {
                return StatusCode(400, ex.Message);
            }
        }
        
        /// <summary>
        /// Method to extract user personality by twitter username
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult AnalyseUserTweets(string username)
        {
            if (username == null || username.Length == 0)
            {
                throw new ArgumentNullException(username, "Twitter Username must be specified");
            }
            Auth.ApplicationCredentials = new TwitterCredentials(CONSUMER_KEY, CONSUMER_SECRET, ACCESS_TOKEN, ACCESS_TOKEN_SECRET);
            var tweets = Timeline.GetUserTimeline(username, 100);
            //Auth.ExecuteOperationWithCredentials(twitterCredentials, () => Timeline.GetUserTimeline(username, 100));

            if (tweets == null)
            {
                throw new Exception("Twitter auth failed");
            }

            return AnalyseData(tweets.Select(x => x.FullText).ToString());
        }

        private JsonResult AnalyseData(string data)
        {
            if (data.Length < 100)
            {
                throw new ArgumentException($"The number of words {data.Length} is less than the minimum number of words required for analysis: 100");
            }

            ContentListContainer contentListContainer = new ContentListContainer
            {
                ContentItems = new List<ContentItem>()
                {
                    new ContentItem()
                    {
                        Contenttype = ContentItem.ContenttypeEnum.TEXT_PLAIN,
                        Language = ContentItem.LanguageEnum.EN,
                        Content = data
                    }
                }
            };

            var result = _personalityInsights.Profile("text/plain", "application/json", contentListContainer, rawScores: true, consumptionPreferences:true, csvHeaders:true);

            return Json(result);
        }

        private string ReadFileContent(IFormFile file)
        {
            var filePath = Path.GetTempFileName();
            var streamFile = file.OpenReadStream();

            switch (file.ContentType)
            {
                case MIMEType.Doc:
                case MIMEType.ODT:
                case MIMEType.Txt:
                    using (var workBook = new WordDocument(streamFile, FormatType.Automatic))
                    {
                        return workBook.GetText();
                    }
                default:
                    throw new ArgumentOutOfRangeException(file.FileName, "File Extension Not allowed");
            }
        }
    }
}
