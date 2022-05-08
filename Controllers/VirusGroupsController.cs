#nullable disable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using LabaOne;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;

namespace LabaOne.Controllers
{
    //[Authorize(Roles="admin, user")]
    public class VirusGroupsController : Controller
    {
        private readonly DBFinalContext _context;

        public VirusGroupsController(DBFinalContext context)
        {
            _context = context;
        }

        // GET: VirusGroups
        public async Task<IActionResult> Index()
        {
            return View(await _context.VirusGroups.ToListAsync());
        }

        // GET: viruses/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
                return NotFound();

            var virus = await _context.VirusGroups
                .FirstOrDefaultAsync(m => m.Id == id);
            if (virus == null)
                return NotFound();
            return RedirectToAction("Index", "Viruses", new { id = virus.Id, name = virus.GroupName });
            //return RedirectToAction("Index", "VirusGroups", new { id = virus.Id, name = virus.GroupName });
        }

        // GET: viruses/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: viruses/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id, GroupName, GroupInfo, DateDiscovered")] VirusGroup group)
        {
            if (ModelState.IsValid)
            {
                _context.Add(group);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(group);
        }

        // GET: viruses/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
                return NotFound();

            var group = await _context.VirusGroups.FindAsync(id);
            if (group == null)
            {
                return NotFound();
            }
            return View(group);
        }

        // POST: Publishers/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id, GroupName, GroupInfo, DateDiscovered")] VirusGroup group)
        {
            if (id != group.Id)
                return NotFound();

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(group);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!PublisherExists(group.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(group);
        }

        // GET: viruses/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var publisher = await _context.VirusGroups
                .FirstOrDefaultAsync(m => m.Id == id);
            if (publisher == null)
            {
                return NotFound();
            }

            return View(publisher);
        }

        // POST: viruses/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var group = await _context.VirusGroups.FindAsync(id);
            var virus = _context.Viruses.Where(c => c.GroupId == id);
            _context.VirusGroups.Remove(group);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool PublisherExists(int id) {
            return _context.VirusGroups.Any(e => e.Id == id);
        }


        






        //public ActionResult Export()
        //{
        //    using (XLWorkbook vir = new XLWorkbook(XLEventTracking.Disabled))
        //    {
        //        var groups = _context.VirusGroups.Include("Viruses").ToList();
        //        //тут, для прикладу ми пишемо усі книжки з БД, в своїх проєктах ТАК НЕ РОБИТИ (писати лише вибрані) 
        //        foreach (var c in groups) {
        //            var worksheet = vir.Worksheets.Add(c.GroupName);

        //            worksheet.Cell("A1").Value = "Назва";
        //            worksheet.Cell("B1").Value = "Вірус 1";
        //            worksheet.Cell("C1").Value = "Вірус 2";
        //            worksheet.Cell("D1").Value = "Вірус 3";
        //            worksheet.Cell("F1").Value = "Група";
        //            worksheet.Cell("G1").Value = "Інформація";
        //            worksheet.Row(1).Style.Font.Bold = true;
        //            var viruses = c.Viruses.ToList();

        //            //нумерація рядків/стовпчиків починається з індекса 1 (не 0) 
        //            for (int i = 0; i < viruses.Count; i++)
        //            {
        //                worksheet.Cell(i + 2, 1).Value = viruses[i].VirusName;
        //                worksheet.Cell(i + 2, 7).Value = viruses[i].Variants;
        //                //var ab = _context.Variants.Where(a => a.VirusId == viruses[i].Id).Include("Author").ToList();

        //                //більше 4-ох нікуди писати 
        //                int j = 0;
        //                //foreach (var a in ab)
        //                //{
        //                //    if (j < 5)
        //                //    {
        //                //        worksheet.Cell(i + 2, j + 2).Value = a.Author.Name;

        //                //        j++;
        //                //    }
        //                //}
        //            }
        //        }
        //        using (var stream = new MemoryStream())
        //        {
        //            vir.SaveAs(stream);
        //            stream.Flush();

        //            return new FileContentResult(stream.ToArray(),
        //                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //            {
        //                //змініть назву файла відповідно до тематики Вашого проєкту 
        //                FileDownloadName = $"library_{DateTime.UtcNow.ToShortDateString()}.xlsx"
        //            };
        //        }
        //    }
        //}










        //[HttpPost]
        //[ValidateAntiForgeryToken]
        //public async Task<IActionResult> Import(IFormFile fileExcel)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        if (fileExcel != null)
        //        {
        //            using (var stream = new FileStream(fileExcel.FileName, FileMode.Create))
        //            {
        //                await fileExcel.CopyToAsync(stream);
        //                using (XLWorkbook workBook = new XLWorkbook(stream, XLEventTracking.Disabled))
        //                {
        //                    var allGroups = (from VirusGroup in _context.VirusGroups
        //                                     select VirusGroup.Na).ToList();

        //                    //перегляд усіх листів (в даному випадку категорій)
        //                    foreach (IXLWorksheet worksheet in workBook.Worksheets)
        //                    {
        //                        //worksheet.Name - назва категорії. Пробуємо знайти в БД, якщо відсутня, то створюємо нову
        //                        VirusGroup newGroup = new VirusGroup();

        //                        var g = (from group in _context.VirusGroups
        //                                 where group.GroupName.Contains(worksheet.Name)
        //                                 select group).ToList();

        //                        if (g.Count > 0)
        //                        {
        //                            continue;
        //                        }
        //                        else
        //                        {
        //                            try
        //                            {
        //                                newGroup.GroupName = worksheet.Name;
        //                                newGroup.GroupInfo = worksheet.Cell("B1").Value.ToString();

        //                                _context.VirusGroups.Add(newGroup);
        //                            }
        //                            catch (Exception ex)
        //                            {
        //                                Console.WriteLine(ex.ToString());
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        await _context.SaveChangesAsync();
        //    }
        //    return RedirectToAction(nameof(Index));
        //}

        //public ActionResult Export()
        //{
        //    using (XLWorkbook workbook = new XLWorkbook(XLEventTracking.Disabled))
        //    {

        //        var groups = _context.VirusGroups.ToList();

        //        foreach (var gr in groups)
        //        {
        //            var worksheet = workbook.Worksheets.Add(gr.GroupName);

        //            worksheet.Cell("A1").Value = "Address";
        //            worksheet.Cell("A2").Value = "PhoneNumber";
        //            worksheet.Cell("A3").Value = "Email";
        //            worksheet.Cell("A4").Value = "Employees";
        //            worksheet.Cell("B4").Value = "Books";
        //            worksheet.Cell("C4").Value = "Orders";

        //            worksheet.Cells("A1:A3").Style.Font.Bold = true;
        //            worksheet.Cells("A4:C4").Style.Font.Bold = true;

        //            worksheet.Column(1).Width = 30;
        //            worksheet.Column(2).Width = 30;
        //            worksheet.Column(3).Width = 6;
        //            worksheet.Column(4).Width = 12;

        //            for (int i = 0, j = 5; i < bookstores.Count; i++)
        //            {
        //                if (bookstores[i].Name == worksheet.Name)
        //                {
        //                    worksheet.Cell(1, 2).Value = bookstores[i].Address;

        //                    worksheet.Cell(2, 2).Value = bookstores[i].PhoneNumber;
        //                    worksheet.Cell(2, 2).SetDataType(XLDataType.Text);

        //                    worksheet.Cell(3, 2).Value = bookstores[i].Email;


        //                    var em = _context.Employees.Where(a => a.IdStore == bookstores[i].Id).ToList();
        //                    foreach (var employee in em)
        //                    {
        //                        worksheet.Cell(j, 1).Value = employee.FullName;
        //                        j++;
        //                    }

        //                    j = 5;
        //                    var b = _context.Books.Where(a => a.IdStore == bookstores[i].Id).ToList();
        //                    foreach (var book in b)
        //                    {
        //                        worksheet.Cell(j, 2).Value = book.Name;
        //                        j++;
        //                    }

        //                    j = 5;
        //                    var orders = (from ord in _context.Orders
        //                                  from emp in _context.Employees
        //                                  where ord.IdEmployee == emp.Id && emp.IdStore == bookstores[i].Id
        //                                  select ord).ToList();

        //                    if (orders.Count() == 0)
        //                    {
        //                        worksheet.Cell(j, 3).Value = "0";
        //                    }
        //                    else
        //                    {
        //                        worksheet.Cell(j, 3).Value = orders.Count();
        //                    }
        //                }
        //                else
        //                {
        //                    continue;
        //                }
        //            }
        //        }

        //        using (var stream = new MemoryStream())
        //        {
        //            workbook.SaveAs(stream);
        //            stream.Flush();
        //            return new FileContentResult(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //            {
        //                FileDownloadName = $"BookStores_{DateTime.UtcNow.ToShortDateString()}.xlsx"
        //            };
        //        }
        //    }
        //}

        //public ActionResult DExport()
        //{
        //    using (MemoryStream ms = new MemoryStream())
        //    {
        //        using (WordprocessingDocument package = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        //        {
        //            var bookstores = _context.BookStores.ToList();

        //            MainDocumentPart mainPart = package.AddMainDocumentPart();
        //            mainPart.Document = new Document();
        //            var body = new Body();

        //            foreach (var store in bookstores)
        //            {
        //                #region RUNs
        //                Run runStoreName = new Run();
        //                Run runAddress = new Run();
        //                Run runPhoneNumber = new Run();
        //                Run runEmail = new Run();

        //                Run runEmployees = new Run();
        //                Run runEmp = new Run();

        //                Run runB = new Run();

        //                Run rubOrd = new Run();

        //                Run run = new Run();

        //                #endregion RUNs

        //                #region PARAGRAPHs
        //                Paragraph storeName = new Paragraph();
        //                Paragraph storeAddress = new Paragraph();
        //                Paragraph storePhoneNumber = new Paragraph();
        //                Paragraph storeEmail = new Paragraph();

        //                Paragraph storeEmployees = new Paragraph();
        //                Paragraph parEmp = new Paragraph();

        //                Paragraph parB = new Paragraph();

        //                Paragraph parOrd = new Paragraph();

        //                Paragraph paragraph = new Paragraph();
        //                #endregion PARAGRAPHs

        //                var em = _context.Employees.Where(a => a.IdStore == store.Id).ToList();
        //                var b = _context.Books.Where(a => a.IdStore == store.Id).ToList();

        //                var orders = (from ord in _context.Orders
        //                              from emp in _context.Employees
        //                              where ord.IdEmployee == emp.Id && emp.IdStore == store.Id
        //                              select ord).ToList();

        //                var d = _context.Deliveries.Where(a => a.Id == store.Id).ToList();

        //                RunProperties runHeaderProperties = runStoreName.AppendChild(new RunProperties(new Bold()));
        //                RunProperties runProperties = runStoreName.AppendChild(new RunProperties(new Italic()));


        //                runStoreName.AppendChild(new Text($"Назва:    		                {store.Name}"));
        //                storeName.Append(runStoreName);
        //                body.Append(storeName);

        //                runAddress.AppendChild(new Text($"Адреса: 		                 {store.Address}"));
        //                storeAddress.Append(runAddress);
        //                body.Append(storeAddress);

        //                runPhoneNumber.AppendChild(new Text($"Номер телефону:  {store.PhoneNumber}"));
        //                storePhoneNumber.Append(runPhoneNumber);
        //                body.Append(storePhoneNumber);

        //                runEmail.AppendChild(new Text($"Email:      	            	    {store.Email}"));
        //                storeEmail.Append(runEmail);
        //                body.Append(storeEmail);

        //                runEmp.AppendChild(new Text("Працівники:"));
        //                parEmp.Append(runEmp);
        //                body.Append(parEmp);

        //                if (em.Count() > 0)
        //                {
        //                    foreach (var employee in em)
        //                    {
        //                        runEmployees.AppendChild(new Text($"•   {employee.FullName}"));
        //                        storeEmployees.Append(runEmployees);
        //                        body.Append(storeEmployees);
        //                    }
        //                }
        //                else
        //                {
        //                    runEmployees.AppendChild(new Text("•   Працівники відсутні."));
        //                    storeEmployees.Append(runEmployees);
        //                    body.Append(storeEmployees);
        //                }

        //                runB.AppendChild(new Text("Книги:"));
        //                parB.Append(runB);
        //                body.Append(parB);


        //                if (b.Count() > 0)
        //                {
        //                    foreach (var book in b)
        //                    {
        //                        Run runBooks = new Run();
        //                        Paragraph storeBooks = new Paragraph();
        //                        runBooks.AppendChild(new Text($"•   {book.Name}"));
        //                        storeBooks.Append(runBooks);
        //                        body.Append(storeBooks);
        //                    }
        //                }
        //                else
        //                {
        //                    Run runBooks = new Run();
        //                    Paragraph storeBooks = new Paragraph();
        //                    runBooks.AppendChild(new Text("•   Книги відсутні."));
        //                    storeBooks.Append(runBooks);
        //                    body.Append(storeBooks);
        //                }

        //                rubOrd.AppendChild(new Text("Замовлення:"));
        //                parOrd.Append(rubOrd);
        //                body.Append(parOrd);


        //                if (orders.Count() != 0)
        //                {
        //                    Run runOrders = new Run();
        //                    Paragraph storeOrders = new Paragraph();
        //                    runOrders.AppendChild(new Text($"•   Кількість замовлень: {orders.Count()}."));
        //                    storeOrders.Append(runOrders);
        //                    body.Append(storeOrders);
        //                }
        //                else
        //                {
        //                    Run runOrders = new Run();
        //                    Paragraph storeOrders = new Paragraph();
        //                    runOrders.AppendChild(new Text("•   Замовлення відсутні."));
        //                    storeOrders.Append(runOrders);
        //                    body.Append(storeOrders);
        //                }

        //                run.AppendChild(new Text("------------------------------------------------------------------------------------------------------------------------------------------"));
        //                paragraph.Append(run);
        //                body.Append(paragraph);
        //            }

        //            mainPart.Document.Append(body);
        //            package.Close();
        //        }

        //        ms.Flush();
        //        return new FileContentResult(ms.ToArray(), "application/vnd.ms-word")
        //        {
        //            //змініть назву файла відповідно до тематики Вашого проєкту
        //            FileDownloadName = $"BookStores_{DateTime.UtcNow.ToShortDateString()}.docx"
        //        };
        //    }
        //}

    }

}
