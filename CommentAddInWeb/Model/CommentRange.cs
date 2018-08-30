using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CommentAddInWeb.Model
{
    public class CommentRange
    {
        public string Id { get; set; }
        public string Initials { get; set; }
        public DateTime? Date { get; set; }
        public string Author { get; set; }
        /// <summary>
        /// Comment的内容
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// 被添加Comment的内容
        /// </summary>
        public string CommentedText { get; set; }
        public int RangeId { get; set; }
        public virtual Range Range { get; set; }
    }
}