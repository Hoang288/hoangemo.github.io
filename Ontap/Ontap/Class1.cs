using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ontap
{
    class Class1
    {
        private string _Hoten;
        private string _Mssv;
        private double _Dqt;
        private double _Dkt;

        public string Hoten
        {
            get
            {
                return _Hoten;
            }
            set
            {
                _Hoten = value;
            }
        }

        public string Mssv
        {
            get
            {
                return _Mssv;
            }
            set
            {
                _Mssv = value;
            }
        }

        public double Dqt
        {
            get
            {
                return _Dqt;
            }
            set
            {
                _Dqt = value;
            }
        }

        public double Dkt
        {
            get
            {
                return _Dkt;
            }
            set
            {
                _Dkt = value;
            }
        }

        public double Tongket()
        {
            return 0.3*Dqt + 0.7*Dkt ;
        }



    }
}
