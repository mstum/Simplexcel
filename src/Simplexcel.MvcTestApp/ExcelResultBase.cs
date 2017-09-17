using System;
using System.IO;
using System.Net;
using System.Web.Mvc;

namespace Simplexcel.MvcTestApp
{
    public abstract class ExcelResultBase : ActionResult
    {
        private readonly string _filename;

        protected ExcelResultBase(string filename)
        {
            _filename = filename;
        }

        protected abstract Workbook GenerateWorkbook();

        public override void ExecuteResult(ControllerContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var workbook = GenerateWorkbook();
            if (workbook == null)
            {
                context.HttpContext.Response.StatusCode = (int)HttpStatusCode.NotFound;
                return;
            }

            context.HttpContext.Response.ContentType = "application/octet-stream";
            context.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
            context.HttpContext.Response.AppendHeader("content-disposition", "attachment; filename=\"" + _filename + "\"");

            workbook.Save(context.HttpContext.Response.OutputStream);
        }
    }
}