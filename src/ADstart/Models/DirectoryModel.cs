using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ADstart.Models
{

    public class DirectoryModel
    {
        /// <summary>
        /// Gets or sets PersonId.
        /// </summary>
        public int UserId { get; set; }


        /// <summary>
        /// Gets or sets OperatorPerson.
        /// </summary>
        public string OperatorPerson { get; set; }

        /// <summary>
        /// Gets or sets Name.
        /// </summary>
        public string Name { get; set; }


        /// <summary>
        /// Gets or sets CN.
        /// </summary>
        public string CN { get; set; }

        /// <summary>
        /// Gets or sets samAccountName.
        /// </summary>
        public string SamAccountName { get; set; }

        /// <summary>
        /// Gets or sets Gender.
        /// </summary>
        public string Gender { get; set; }

        /// <summary>
        /// Gets or sets City.
        /// </summary>
        public string City { get; set; }



        /// OU:
        /// 
        /// <summary>
        /// Gets or sets OU.
        /// </summary>
        public string OU { get; set; }

        /// <summary>
        /// Gets or sets OUDescription.
        /// </summary>
        public string OUDescription { get; set; }

        /// User:
        /// 
        /// <summary>
        /// Gets or sets OU.
        /// </summary>
        public string UserName { get; set; }




    }
}