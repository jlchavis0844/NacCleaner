using System.Text.RegularExpressions;

namespace NacCleaner {
    internal class CLineLife {
        public string name { get; set; }
        public string policyNum { get; set; }
        public string type { get; set; }
        public string plan { get; set; }
        public string accDate { get; set; }
        public string dueDate { get; set; }
        public int mopd { get; set; }
        public double premium { get; set; }
        public double cRate { get; set; }
        public double split { get; set; }
        public double comm { get; set; }

        public CLineLife(string name, string policyNum, string type, string plan, string accDate, string dueDate, int mopd,
                double premium, double cRate, double split, double comm) {
            Regex nameReg = new Regex(@"[^a-zA-Z\s]");
            this.name = nameReg.Replace(name, "").Trim();
            //this.name = name;
            this.policyNum = policyNum;
            this.type = type;
            this.plan = plan;
            this.accDate = accDate;
            this.dueDate = dueDate;
            this.mopd = mopd;
            this.premium = premium;
            this.cRate = cRate;
            this.split = split;
            this.comm = comm;
        }

        public CLineLife() {
            this.name = "";
            this.policyNum = "";
            this.type = "";
            this.plan = "";
            this.accDate = "";
            this.dueDate = "";
            this.mopd = 0;
            this.premium = 0.0;
            this.cRate = 0.0;
            this.split = 0.0;
            this.comm = 0.0;
        }

        public string toString() {
            return name + "\t" + policyNum + "\t" + type + "\t" + plan + "\t" + accDate + "\t" + dueDate +
                    "\t" + mopd + "\t" + premium + "\t" + cRate + "\t" + split + "\t" + comm;
        }
    }
}