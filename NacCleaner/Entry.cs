using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace NacCleaner {
    internal class Entry {
        private string name;
        private string policyNum;
        private double rate;
        private List<CLineLife> CLineLifes;
        private List<ALineLife> ALineLifes;
        private TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

        public Entry(string name, string policyNum, double rate) {
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            this.name = nameReg.Replace(name, "").Trim();
            //this.name = name;
            this.policyNum = policyNum;
            CLineLifes = new List<CLineLife>();
            ALineLifes = new List<ALineLife>();
            this.rate = rate;
        }

        public double getCommissionTotal() {
            double total = 0.0;
            double commissions = 0.0;
            double commApps = 0.0;
            double cBacks = 0.0;
            double advs = 0.0;

            foreach (CLineLife cl in CLineLifes) {
                commissions += cl.comm;
            }

            foreach (ALineLife al in ALineLifes) {
                commApps += al.commApp;
                cBacks += al.chargeBack;
                advs += al.currAdv;
            }

            total = commissions - commApps + cBacks + advs;
            return total;
        }

        public string printOut() {
            string strOut = "";

            strOut = policyNum + ", " + name + ", " + CLineLifes[0].plan + ", " + CLineLifes[0].accDate + ", " +
            CLineLifes[0].premium + ", " + rate + ", ";

            if (CLineLifes[0].split < 0.1) {
                strOut += "100,";
            }
            else {
                strOut += (CLineLifes[0].split * 100) + ", ";
            }

            if (CLineLifes[0].type == "RN") {
                strOut += "0, " + getCommissionTotal();
            }
            else {
                strOut += getCommissionTotal() + ", 0";
            }
            Console.WriteLine(strOut);
            return strOut;
        }

        public Object[] getOutput() {
            Object[] ret = new Object[9];
            ret[0] = policyNum;
            ret[1] = textInfo.ToTitleCase(name.Replace("*", "").ToLower());
            ret[2] = CLineLifes[0].plan;
            ret[3] = CLineLifes[0].accDate;
            ret[4] = CLineLifes[0].premium;
            ret[5] = CLineLifes[0].cRate * 100;
            ret[6] = CLineLifes[0].split;

            if (CLineLifes[0].type == "RN") {
                ret[8] = getCommissionTotal();
                ret[7] = 0.0;
            }
            else {
                ret[7] = getCommissionTotal();
                ret[8] = 0.0;
            }
            return ret;
        }

        public double getPremium() {
            return CLineLifes[0].premium;
        }

        public double getRatePer() {
            return CLineLifes[0].cRate;
        }

        public double getSplit() {
            double tSplit = CLineLifes[0].split;
            if (tSplit == 0.0) {
                return 100.0;
            }
            else return tSplit * 100;
        }

        public string getType() {
            return CLineLifes[0].type;
        }


        /**
         * @return the rate
         */
        public double getRate() {
            return rate * 100;
        }


        /**
         * @param rate the rate to set
         */
        public void setRate(double rate) {
            this.rate = rate;
        }


        public void addCLineLife(CLineLife cl) {
            CLineLifes.Add(cl);
        }

        public void addALineLife(ALineLife al) {
            ALineLifes.Add(al);
        }

        public string getPlan() {
            return CLineLifes[0].plan;
        }

        public string getIssueDate() {
            return CLineLifes[0].accDate;
        }

        public void setIssueDate(string date) {
            CLineLifes[0].accDate = date;
        }

        /**
         * @return the name
         */
        public string getName() {
            return name;
        }

        /**
         * @param name the name to set
         */
        public void setName(string name) {
            this.name = name;
        }

        /**
         * @return the policyNum
         */
        public string getPolicyNum() {
            return policyNum;
        }

        /**
         * @param policyNum the policyNum to set
         */
        public void setPolicyNum(string policyNum) {
            this.policyNum = policyNum;
        }


        /**
         * @return the CLineLifes
         */
        public List<CLineLife> getCLineLifes() {
            return CLineLifes;
        }

        /**
         * @param CLineLifes the CLineLifes to set
         */
        public void setCLineLifes(List<CLineLife> CLineLifes) {
            this.CLineLifes = CLineLifes;
        }

        /**
         * @return the ALineLifes
         */
        public List<ALineLife> getALineLifes() {
            return ALineLifes;
        }

        /**
         * @param ALineLifes the ALineLifes to set
         */
        public void setALineLifes(List<ALineLife> ALineLifes) {
            this.ALineLifes = ALineLifes;
        }
    }
}