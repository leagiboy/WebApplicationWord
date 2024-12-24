using System.ComponentModel.DataAnnotations;

namespace WebApplicationWord.Models
{
    public class Student
    {
        /// <summary>
        /// ID 主键
        /// </summary>
        [Key]
        public int Id { get; set; }

        [Required, StringLength(30)]
        public string StudentsName { get; set; }

        [Required, StringLength(18), Length(18, 18)]
        public string IdCardNo { get; set; }

        //根据身份证自动计算年龄
        public int? Age => DateTime.Today.Year - int.Parse(IdCardNo.Substring(6, 4));
    }
}