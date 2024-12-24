using System.Data;
using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using NPOI.XWPF.UserModel;
using WebApplicationWord.Models;

namespace WebApplicationWord.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly AppDbcontext _appDbcontext;

        public HomeController(ILogger<HomeController> logger, AppDbcontext context)
        {
            _logger = logger;
            _appDbcontext = context;
        }

        public IActionResult Index()
        {
            var students = _appDbcontext.Students.ToList();

            return View(students);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        /// <summary>
        /// 使用word格式导出数据
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportWord()
        {
            var students = _appDbcontext.Students.ToList();
            //创建word对象
            XWPFDocument document = new XWPFDocument();
            //创段落和表格
            XWPFParagraph para = document.CreateParagraph();
            XWPFRun run = para.CreateRun();
            run.SetText("学生清单");
            run.IsBold = true;
            run.FontSize = 16;
            //创建表格
            XWPFTable table = document.CreateTable();
            XWPFTableRow headrow = table.GetRow(0);
            headrow ??= table.CreateRow();
            headrow.GetCell(0).SetText("ID");
            headrow.AddNewTableCell().SetText("姓名");
            headrow.AddNewTableCell().SetText("身份证号");
            headrow.AddNewTableCell().SetText("年龄");
            foreach (var student in students)
            {
                XWPFTableRow datarow = table.CreateRow();
                datarow.GetCell(0).SetText(student.Id.ToString());
                datarow.GetCell(1).SetText(student.StudentsName);
                datarow.GetCell(2).SetText(student.IdCardNo);
                datarow.GetCell(3).SetText(student.Age.ToString());
            }
            MemoryStream stream = new MemoryStream();
            document.Write(stream);
            string filename = $"学生清单{DateTime.Now.ToString("yyyyMMdd")}.docx";
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename);
        }
    }
}