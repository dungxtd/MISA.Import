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
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MISA.Import.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        public static List<Customer> customers = new List<Customer>();
        private String connectionString = "" +
        "Host = 47.241.69.179;" +
        "Port = 3306;" +
        "Database = MF826_Import_TDDUNG;" +
        "User Id = dev;" +
        "Password = 12345678;" +
        "convert zero datetime=true";

        private IDbConnection dbConnection;
        private string checkString(Object obj)
        {
            if (obj == null)
                return null;
            else return (obj ?? string.Empty).ToString().Trim();
        }
        private string checkStringUTF8(Object obj)
        {
            if (obj == null)
                return null;
            else return (obj ?? string.Empty).ToString().Trim();
        }
        private Nullable<DateTime> checkDateTime(Object obj)
        {
            if (obj == null)
                return null;
            else
            {
                String dateTime = (obj ?? string.Empty).ToString().Trim();
                if (dateTime.Length == 10) return DateTime.ParseExact( dateTime, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if (dateTime.Length == 7|| dateTime.Length == 6) return DateTime.ParseExact(dateTime, "MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if (dateTime.Length == 4) return DateTime.ParseExact(dateTime, "yyyy", System.Globalization.CultureInfo.InvariantCulture);

            }
            return null;
        }

        private bool CheckCustomerCode(string customerCode)
        {
            using (dbConnection = new MySqlConnection(connectionString))
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@m_CustomerCode", customerCode);
                var rowsEffect = dbConnection.QueryFirstOrDefault<bool>("CheckCustomerCodeExists", param: dynamicParameters, commandType: CommandType.StoredProcedure);
                return rowsEffect;
            }
        }
        private bool CheckPhoneNumber(string phoneNumber)
        {

            using (dbConnection = new MySqlConnection(connectionString))
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@m_PhoneNumber", phoneNumber);
                var rowsEffect = dbConnection.QueryFirstOrDefault<bool>("CheckPhoneNumberExists", param: dynamicParameters, commandType: CommandType.StoredProcedure);
                return rowsEffect;
            }
        }
        private Guid? GetCustomerGroupIdByName(byte[] customerGroupName)
        {
            var utf8_customerGroupName = Encoding.UTF8.GetString(customerGroupName);
            using (dbConnection = new MySqlConnection(connectionString))
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                dynamicParameters.Add("@m_CustomerGroupName", utf8_customerGroupName);
                var customerGroupid = dbConnection.QueryFirstOrDefault<Guid>("GetCustomerGroupIdByName", param: dynamicParameters, commandType: CommandType.StoredProcedure);
                if (customerGroupid == new Guid("00000000-0000-0000-0000-000000000000")) return null;
                return customerGroupid;
            }
        }

        private void GetDataFromFile(List<Customer> list)
        {
            //check customer
            for (int i = 0; i < list.Count; i++)
            {
                list[i].Status = null;
                int[] tempError = { 0, 0, 0, 0, 0 };
                for (int j = 0; j < i; j++)
                {
                    if (list[j].CustomerCode.Equals(list[i].CustomerCode)) tempError[0] = 1;
                    if (list[j].PhoneNumber.Equals(list[i].PhoneNumber)) tempError[1] = 1;
                }
                if (CheckCustomerCode(list[i].CustomerCode)) tempError[2] = 1;
                if (CheckPhoneNumber(list[i].CustomerCode)) tempError[3] = 1;
                if (list[i].CustomerGroupId == null) tempError[4] = 1;
                if (tempError[0] == 1) list[i].Status += "Mã khách hàng đã trùng với khách hàng khác trong tệp nhập khẩu.";
                if (tempError[1] == 1) list[i].Status += "SĐT đã trùng với SĐT của khách hàng khác trong tệp nhập khẩu.";
                if (tempError[2] == 1) list[i].Status += "Mã khách hàng đã tồn tại trong hệ thống.";
                if (tempError[3] == 1) list[i].Status += "SĐT đã có trong hệ thống.";
                if (tempError[4] == 1) list[i].Status += "Nhóm khách hàng không có trong hệ thống.";
            }
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].Status == null)
                {
                    list[i].Status = "Hợp lệ.";
                    list[i].StatusCode = 1;
                }
                else list[i].StatusCode = 0;
            }
        }



        [HttpPost]
        public async Task<List<Customer>> ImportAsync(IFormFile formFile, CancellationToken cancellationToken)
        {
            if (formFile == null || formFile.Length <= 0)
            {
                throw new BadRequestException("formfile is empty");
            }

            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new BadRequestException("Not Support file extension");
            }
            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        customers.Add(new Customer
                        {
                            CustomerCode = checkString(worksheet.Cells[row, 1].Value),
                            Fullname = checkString(worksheet.Cells[row, 2].Value),
                            MemberCardCode = checkString(worksheet.Cells[row, 3].Value),
                            CustomerGroupId = GetCustomerGroupIdByName(Encoding.Default.GetBytes(checkString(worksheet.Cells[row, 4].Value))),
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
            GetDataFromFile(customers);
            return customers;
        }


        [HttpPost("AddToDb")]
        public ActionResult ImportToDb()
        {

            //var rowsAffect = 0;
            var temp = 0;
            for (int i = 0; i < customers.Count; i++)
            {
                DynamicParameters dynamicParameters = new DynamicParameters();
                if (customers[i].StatusCode == 1)
                {
                    dynamicParameters.Add("@Customer", customers[i]);
                    using (dbConnection = new MySqlConnection(connectionString))
                    {
                        var rowsAffect = dbConnection.Execute("InsertCustomer", param: customers[i], commandType: CommandType.StoredProcedure);
                        if (rowsAffect != 1) return BadRequest("Có lỗi xảy ra vui lòng kiểm tra lại.");
                        else temp = 1;
                    }
                }
                customers[i].Status = null;
            }
            if (temp == 1) {
                temp = 0;
                customers.Clear();
                return Ok();
            } 
            else return BadRequest("Có lỗi xảy ra vui lòng kiểm tra lại.");
        }
    }
}
