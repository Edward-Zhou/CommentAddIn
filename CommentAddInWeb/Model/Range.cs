using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CommentAddInWeb.Model
{
    public class Range
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public virtual ICollection<CommentRange> CommentRanges { get; set; }
    }
}