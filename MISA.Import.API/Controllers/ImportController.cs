using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Core.Exceptions;
using MISA.Import.Core.Entities;
using MySqlConnector;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MISA.Import.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        private string checkString(Object obj)
        {
            if (obj == null)
                return null;
            else return (obj ?? string.Empty).ToString().Trim();
        }


        private DateTime checkDateTime(Object obj)
        {
            if (obj == null)
                return DateTime.MinValue;
            else return DateTime.ParseExact((obj ?? string.Empty).ToString().Trim(), "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
        }


        [HttpPost("import")]
        public async Task<int> Import(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                throw new BadRequestException("formfile is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new BadRequestException("Not Support file extension");
            }

            var list = new List<Customer>();

            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 3; row <= rowCount; row++)
                    {
                        list.Add(new Customer
                        {
                            CustomerCode = checkString(worksheet.Cells[row, 1].Value),
                            Fullname = checkString(worksheet.Cells[row, 2].Value),
                            MemberCardCode = checkString(worksheet.Cells[row, 3].Value),
                            CustomerGroup = checkString(worksheet.Cells[row, 4].Value),
                            PhoneNumber = checkString(worksheet.Cells[row, 5].Value),
                            DateOfBirth = checkDateTime(worksheet.Cells[row, 6].Value),
                            CompanyName = checkString(worksheet.Cells[row, 7].Value),
                            TaxCode = checkString(worksheet.Cells[row, 8].Value),
                            Email = checkString(worksheet.Cells[row, 9].Value),
                            Address = checkString(worksheet.Cells[row, 10].Value),
                            Note = checkString(worksheet.Cells[row, 11].Value),
                        });
                    }
                }
            }


            String connectionString = "" +
                "Host = 47.241.69.179;" +
                "Port = 3306;" +
                "Database = MF826_Import_TDDUNG;" +
                "User Id = dev;" +
                "Password = 12345678;" +
                "convert zero datetime=true";

            IDbConnection dbConnection;
            var rowsAffect = 894;
            // add list to db ..  
            foreach (Customer customer in list)
            {
                using (dbConnection = new MySqlConnection(connectionString))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    dynamicParameters.Add($"Customer", customer);
                    rowsAffect = dbConnection.Execute("InsertCustomer", param: customer, commandType: CommandType.StoredProcedure);
                }
            }

            // here just read and return  

            return rowsAffect;

        }
         



    }
}
