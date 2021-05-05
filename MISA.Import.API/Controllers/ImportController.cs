using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Import.Core.Entities;
using MISA.Import.Core.Responses;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MISA.Import.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        [HttpPost("import")]
        public async Task<CustomerResponse<List<Customer>>> Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                return CustomerResponse<List<Customer>>.GetResult(-1, "formfile is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return CustomerResponse<List<Customer>>.GetResult(-1, "Not Support file extension");
            }

            var list = new List<Customer>();
            

            using (var stream = new MemoryStream())
            {

                await formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    for (int row = 3; row <= rowCount; row++)
                    {   
                        for (int col = 1; row <= colCount; col++)
                        {
                            //if (worksheet.Cells[row, col] == null) 
                        }
                            var check = worksheet.Cells[4, 9].Value;
                        if (check != null)
                        {
                            check = check.ToString().Trim();
                            list.Add(new Customer
                            {
                                CustomerCode = worksheet.Cells[row, 1].Value.ToString().Trim(),
                                Fullname = worksheet.Cells[row, 2].Value.ToString().Trim(),
                                MemberCardCode = worksheet.Cells[row, 3].Value.ToString().Trim(),
                                CustomerGroup = worksheet.Cells[row, 4].Value.ToString().Trim(),
                                PhoneNumber = worksheet.Cells[row, 5].Value.ToString().Trim(),
                                CompanyName = worksheet.Cells[row, 7].Value.ToString().Trim(),
                                TaxCode = worksheet.Cells[row, 8].Value.ToString().Trim(),
                                Email = null,
                                Address = worksheet.Cells[row, 10].Value.ToString().Trim(),
                                Note = worksheet.Cells[row, 11].Value.ToString().Trim(),
                            });
                        }
                    }
                }
            }

            //DateOfBirth = (worksheet.Cells[row, 6].Value.ToString().Trim()),
            //CompanyName = worksheet.Cells[row, 7].Value.ToString().Trim(),
            //TaxCode = worksheet.Cells[row, 8].Value.ToString().Trim(),
            //Email = worksheet.Cells[row, 9].Value.ToString().Trim(),
            //Address = worksheet.Cells[row, 10].Value.ToString().Trim(),
            //Note = worksheet.Cells[row, 11].Value.ToString().Trim(),
            // add list to db ..  
            // here just read and return  

            return CustomerResponse<List<Customer>>.GetResult(0, "OK", list);
        }
    }
}
