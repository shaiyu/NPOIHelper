using System;
using System.Buffers;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AspNetCoreSample.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOIHelper;
using NPOIHelperWebExtension;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace AspNetCoreSample.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ImportExcelController : ControllerBase
    {
        [HttpPost]
        public string Import()
        {
            var file = Request.Form.Files[0];
            var contentType = file.ContentType;
            using (var stream = file.OpenReadStream())
            {
                stream.Seek(0, SeekOrigin.Begin);
                var reader = NPOIHelper.NPOIHelperBuild.GetReader(stream, contentType);
                var schoolRecommends = reader.ReadSheet<ImportDto>(sheetIndex: 0).ToList();
            }
            return "success";
        }


        [HttpPost]
        public string Import2()
        {
            var reader = NPOIHelper.NPOIHelperBuild.GetReader("Temp/8c5ca2c1-58ae-4679-8a18-7f8784265853.xlsx");
            var schoolRecommends = reader.ReadSheet<ImportDto>(sheetIndex: 0).ToList();
            return "success";
        }
    }
}
