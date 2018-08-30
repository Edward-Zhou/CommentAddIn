using CommentAddInWeb.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace CommentAddInWeb
{
    [RoutePrefix("api/comments")]
    public class CommentsController : ApiController
    {
        [Route("ConvertOOXmlToComments")]
        [HttpPost]
        public IEnumerable<CommentRange> ConvertOOXmlToComments([FromBody] string xml)
        {
            var stream = OOXml.GetPackageStreamFromWordOpenXML(xml);
            var comments = OOXml.ConvertStreamToCommentRange(stream);
            return comments;
        }
    }
}
