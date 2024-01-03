[HttpPost]

        public IActionResult DownloadDataMixingExcel([FromForm]int version)

        {

            var claims = hostingDictionary.HttpContext.Current.User.Claims.ToList();

            string username = claims[0].Value;

            try

            {

 

                //println

                var workbook = new XLWorkbook();

                workbook.AddWorksheet("Veri Karıştırma");

                var ws = workbook.Worksheet("Veri Karıştırma");

                ws.Cells("A1").Style.Font.Bold = true;

                ws.Cell("A1").Value = "Veri No";

                ws.Cells("B1").Style.Font.Bold = true;

                ws.Cell("B1").Value = "Üst Veri Grubu";

                ws.Cells("C1").Style.Font.Bold = true;

                ws.Cell("C1").Value = "Veri Grubu";

                ws.Cells("D1").Style.Font.Bold = true;

                ws.Cell("D1").Value = "Veri Adı";

                ws.Cells("E1").Style.Font.Bold = true;

                ws.Cell("E1").Value = "Veri Tanımı";

                ws.Cells("F1").Style.Font.Bold = true;

                ws.Cell("F1").Value = "Veri Sahibi(Owner)";

                ws.Cells("G1").Style.Font.Bold = true;

                ws.Cell("G1").Value = "Veri Bakıcısı(Custodian)";

                ws.Cells("H1").Style.Font.Bold = true;

                ws.Cell("H1").Value = "Veri Kullanıcısı(User)";

                ws.Cells("I1").Style.Font.Bold = true;

                ws.Cell("I1").Value = "Gizlilik Değeri";

                ws.Cells("J1").Style.Font.Bold = true;

                ws.Cell("J1").Value = "Nedeni(Gizli ve Çok Gizli İçin)";

                ws.Cells("K1").Style.Font.Bold = true;

                ws.Cell("K1").Value = "Kapsam";

                ws.Cells("L1").Style.Font.Bold = true;

                ws.Cell("L1").Value = "Karıştırma Kuralı";

                ws.Cells("M1").Style.Font.Bold = true;

                ws.Cell("M1").Value = "Karıştırma Açıklaması";

 

                int row = 2;

                var result = _datamixing.DataMixingExcelModel(version);

                foreach (var item in result)

                {

                    ws.Cell("A" + row.ToString()).Value = item.DataDictionaryId.ToString();

                    ws.Cell("B" + row.ToString()).Value = (item.ParentDataGroupName == null) ? "" : item.ParentDataGroupName.ToString();

                    ws.Cell("C" + row.ToString()).Value = (item.DataGroupName == null) ? "" : item.DataGroupName.ToString();

                    ws.Cell("D" + row.ToString()).Value = (item.DataName == null) ? "" : item.DataName.ToString();

                    ws.Cell("E" + row.ToString()).Value = (item.DataDescription == null) ? "" : item.DataDescription.ToString();

                    ws.Cell("F" + row.ToString()).Value = (item.DataOwnerName == null) ? "" : item.DataOwnerName.ToString();

                    ws.Cell("G" + row.ToString()).Value = (item.DataCustodianName == null) ? "" : item.DataCustodianName.ToString();

                    ws.Cell("H" + row.ToString()).Value = (item.DepartmentName == null) ? "" : item.DepartmentName.ToString();

                    ws.Cell("I" + row.ToString()).Value = (item.DataPrivacyId == null) ? "" : item.DataPrivacyId.ToString();

                    ws.Cell("J" + row.ToString()).Value = (item.DataPrivacyReason == null) ? "" : item.DataPrivacyReason.ToString();

                    ws.Cell("K" + row.ToString()).Value = (item.DataMixingScope == null) ? "" : item.DataMixingScope.ToString();

                    ws.Cell("L" + row.ToString()).Value = (item.DataMixingRule == null) ? "" : item.DataMixingRule.ToString();

                    ws.Cell("M" + row.ToString()).Value = (item.DataMixingRuleDescription == null) ? "" : item.DataMixingRuleDescription.ToString();

 

                    row++;

 

                }

 

                var roothPath = Path.Combine(_environment.ContentRootPath, "download-data-excel");

                if (!Directory.Exists(roothPath))

                    Directory.CreateDirectory(roothPath);

 

 

                string fileName = "VKC" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

                var filePath = Path.Combine(roothPath, fileName);

                if (Directory.Exists(roothPath))

                {

                    string[] files = Directory.GetFiles(roothPath);

                    foreach (string f in files)

                    {

                        System.IO.File.Delete(f);

                   }

                }

 

 

                using (var stream = new FileStream(filePath, FileMode.Create))

                {

 

 

                    workbook.SaveAs(stream);

                    _EahouseContext.SaveChanges();

 

 

 

 

                }

 

                var provider = new FileExtensionContentTypeProvider();

                var file = Path.Combine(_environment.ContentRootPath, "download-data-excel", fileName);

                string contentType;

                if (!provider.TryGetContentType(file, out contentType))

                {

                    contentType = "application/octet-stream";

                }

                byte[] fileBytes;

                if (System.IO.File.Exists(file))

                {

 

                    fileBytes = System.IO.File.ReadAllBytes(file);

                }

                else

                {

                    return NotFound();

                }

                Response.Headers.Add("Content-Disposition", $"attachment;filename=\"{fileName}\"");

                MemoryStream ms = new MemoryStream(fileBytes);

                return new FileStreamResult(ms, contentType);

            }

            catch (Exception e)

            {

                var model = new AppLog()

                {

                    LogMessages = e.Message,

                    UserName = username,

                    Date = DateTime.Now

                };

 

 

                _EahouseContext.AppLogs.Add(model);

                _EahouseContext.SaveChanges();

                return StatusCode(500, e.Message + e.StackTrace);

            }

        }