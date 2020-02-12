using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace probanr2
{
    class Customer
    {
        protected string name, address, email, county;
        protected double phone, reference;

        public Customer(string name, string address = "none", string email = "none", string county = "none", 
            double phone = 1, double reference = 1)
        {
            this.name = name;
            this.address = address;
            this.email = email;
            this.county = county;
            this.phone = phone;
            this.reference = reference;
        }

        public string getName()
        {
            return name;
        }

        public string getAddress()
        {
            return address;
        }

        public string getEmail()
        {
            return email;
        }

        public string getCounty()
        {
            return county;
        }

        public double getPhone()
        {
            return phone;
        }

        public double getReference()
        {
            return reference;
        }

    }
}
