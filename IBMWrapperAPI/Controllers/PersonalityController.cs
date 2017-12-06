using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IBM.WatsonDeveloperCloud.PersonalityInsights.v3;
using IBM.WatsonDeveloperCloud.PersonalityInsights.v3.Model;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using System;
using System.Collections.Generic;
using System.Linq;

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

        /// <summary>
        /// Wrapper Controller Constructor
        /// </summary>
        public PersonalityController()
        {
            IBMId = "f7de21f9-1eee-4e47-9ff2-1040d6c762cb";
            IBMPassword = "WJZs7TfsDzdI";
            _personalityInsights = new PersonalityInsightsService(IBMId, IBMPassword, "2017-10-13");
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
        public IActionResult AnalyseFile([FromBody] IEnumerable<IFormFile> files)
        {
            try
            {
                var file = files.First();
                if (file == null)
                {
                    throw new ArgumentException("File is required");
                }

                var data = ReadExcelFile(file);

                return AnalyseData(data);
            }
            catch (Exception ex)
            {
                return StatusCode(400, ex.Message);
            }
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

        private string ReadExcelFile(IFormFile file)
        {
            var x = file.OpenReadStream();

            using (var spreadsheet = SpreadsheetDocument.Open(x, false))
            {
                var content = default(string);

                var workBookPart = spreadsheet.WorkbookPart;
                var workSheetPart = workBookPart.WorksheetParts.FirstOrDefault();

                var reader = OpenXmlReader.Create(workSheetPart);

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        content = $"{content} {reader.GetText()}";
                    }
                }

                return content;
            }
        }
    }
}
