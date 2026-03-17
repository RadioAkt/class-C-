using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace задание_1
{
    internal class Student
    {
        private string surname;
        private int mat;
        private int fiz;
        private int him;
        private int inf;
        private string citizenship;
        public string Surname
        {
            set
            {
                surname = value;
            }
            get
            {
                return surname; 
            }

        }
        public int Mat
        {
            set
            {
                mat = value;
            }
            get
            {
                return mat;
            }
        }
        public int Fiz
        {
            set
            {
                fiz = value;
            }
            get
            {
                return fiz;
            }
        }
        public int Him
        {
            set
            {
                him = value; 
            }
            get
            {
                return him;
            }
        }
        public int Inf
        {
            set
            {
                inf = value;
            }
            get
            {
                return inf;
            }
        }
        public string Citizenship
        {
            set
            {
                citizenship = value;
            }
            get
            {
                return citizenship;
            }
        }
        public int Stipa(int basest)
        {
            int[] score = {mat, fiz, him, inf};
            int s2 = 0;
            int s3 = 0;
            int s4 = 0;
            int s5 = 0;
            for (int i = 0; i < score.Length; i++)
            {
                if (score[i] == 2)
                    s2++;
                else if (score[i] == 3)
                    s3++;
                else if (score[i] == 4)
                    s4++;
                else
                    s5++;
            }
            int stip = basest;
            if (s2 > 0)
            {
                stip = 0;
            }
            else if (s2 == 0 & s3 != 0)
            {
                if (citizenship != "Россия")
                {
                    stip = basest;
                }
                else
                    stip = basest * 6 / 5;
            }
            else if (s3 == 0 & s4 != 0)
            {
                if (s4 == 4 & citizenship != "Россия" | s5 == 1 & citizenship != "Россия" & s4 == 3)
                {
                    stip = basest * 6 / 5;
                }
                else if (s4 == 4 & citizenship == "Россия" | s5 == 1 & citizenship == "Россия" & s4 == 3)
                {
                    stip = basest * 7 / 5;
                }
                else if (s5 >= 2 & s4 != 0 & citizenship != "Россия")
                {
                    stip = basest * 5 / 4;
                }
                else
                {
                    stip = basest * 29 / 20;
                }
            }
            else if (s5 == 4 & citizenship == "США")
            {
                stip = basest * 5 / 4;
            }
            return stip;
        }
    }
}
