using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MISA.Import.Core.Entities
{
    public class Customer
    {
        ///Thông tin khách hàng
        /// CreatedBy: TDDUNG (05/05/2021)
        ///<summary>
        ///Khoá chính
        /// </summary>
        public Guid CustomerId { get; set; }

        ///<summary>
        ///Mã khsach hàng
        /// </summary>
        public string CustomerCode { get; set; }

        ///<summary>
        ///Họ và tên
        /// </summary>
        public string Fullname { get; set; }

        ///<summary>
        ///Mã thẻ thành viên
        /// </summary>
        public string MemberCardCode { get; set; }

        ///<summary>
        ///Nhóm khách hàng
        /// </summary>
        public string CustomerGroup { get; set; }

        ///<summary>
        ///Số điện thoại
        /// </summary>
        public string PhoneNumber { get; set; }

        ///<summary>
        ///Ngày sinh
        /// </summary>
        public DateTime DateOfBirth { get; set; }

        ///<summary>
        ///Tên công ty
        /// </summary>
        public string CompanyName { get; set; }

        ///<summary>
        ///Mã tax công ty
        /// </summary>
        public string TaxCode { get; set; }

        ///<summary>
        ///Email
        /// </summary>
        public string Email { get; set; }

        ///<summary>
        ///Địa chỉ
        /// </summary>
        public string Address { get; set; }

        ///<summary>
        ///Ghi chú
        /// </summary>
        public string Note { get; set; }

        ///<summary>
        ///Ngày tạo
        /// </summary>
        public DateTime? CreatedDate { get; set; }

        ///<summary>
        ///Tạo bởi
        /// </summary>
        public string CreatedBy { get; set; }

        ///<summary>
        ///Ngày sửa đổi
        /// </summary>
        public DateTime ModifiedDate { get; set; }

        ///<summary>
        ///Sửa bởi
        /// </summary>
        public string ModifiedBy { get; set; }
    }
}
