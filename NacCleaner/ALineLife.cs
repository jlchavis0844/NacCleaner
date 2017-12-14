using System.Text.RegularExpressions;

namespace NacCleaner {
    internal class ALineLife {
        public string name { get; set; }
        public string policyNum { get; set; }
        public string accDate { get; set; }
        public string issueDate { get; set; }
        public int mopd { get; set; }
        public double beginBal { get; set; }
        public double currAdv { get; set; }
        public double commApp { get; set; }
        public double chargeBack { get; set; }
        public double endBal { get; set; }

        public ALineLife(string name, string policyNum, string accDate, string issueDate, int mopd, double beginBal,
            double currAdv, double commApp, double chargeBack, double endBal) {
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            this.name = nameReg.Replace(name, "").Trim();
            //this.name = name;
            this.policyNum = policyNum;
            this.accDate = accDate;
            this.issueDate = issueDate;
            this.mopd = mopd;
            this.beginBal = beginBal;
            this.currAdv = currAdv;
            this.commApp = commApp;
            this.chargeBack = chargeBack;
            this.endBal = endBal;
        }

        public ALineLife() {
            this.name = "";
            this.policyNum = "";
            this.accDate = "";
            this.issueDate = "";
            this.mopd = 0;
            this.beginBal = 0.0;
            this.currAdv = 0.0;
            this.commApp = 0.0;
            this.chargeBack = 0.0;
            this.endBal = 0.0;
        }

        public override string ToString() {
            return name + "\t" + policyNum + "\t" + accDate + "\t" + issueDate + "\t" + mopd + "\t" + beginBal + "\t" + currAdv + "\t" +
                    commApp + "\t" + chargeBack + "\t" + endBal;
        }
    }
}