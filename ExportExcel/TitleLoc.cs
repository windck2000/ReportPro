namespace ExportExcel
{
    // Token: 0x02000004 RID: 4
    public class TitleLoc
    {
        // Token: 0x06000001 RID: 1 RVA: 0x00002050 File Offset: 0x00000250
        public TitleLoc(int snCount)
        {
            this.DataStartLOC = 6;
            this.DataEndLOC = this.DataStartLOC + snCount - 1;
            this.MAX_SPC += snCount;
            this.MIN_SPC += snCount;
            this.UNIT += snCount;
            this.MAX += snCount;
            this.MIN += snCount;
            this.AVG += snCount;
            this.STD += snCount;
            this.Cpu += snCount;
            this.Cpl += snCount;
            this.Cp_M_1 += snCount;
            this.Ca_L_1 += snCount;
            this.Cpk_M_1 += snCount;
            this.Result += snCount;
        }

        // Token: 0x06000002 RID: 2 RVA: 0x00002198 File Offset: 0x00000398
        public string getCellFormula(TitleTYpe FormulaType, int colNO)
        {
            string result = string.Empty;
            string text = TitleLoc.Number2ExcelCol(colNO);
            switch (FormulaType)
            {
                case TitleTYpe.MAX:
                    result = string.Format("MAX({0}{1}:{0}{2})", text, this.DataStartLOC.ToString(), this.DataEndLOC.ToString());
                    break;
                case TitleTYpe.MIN:
                    result = string.Format("MIN({0}{1}:{0}{2})", text, this.DataStartLOC.ToString(), this.DataEndLOC.ToString());
                    break;
                case TitleTYpe.AVG:
                    result = string.Format("AVERAGE({0}{1}:{0}{2})", text, this.DataStartLOC.ToString(), this.DataEndLOC.ToString());
                    break;
                case TitleTYpe.STD:
                    result = string.Format("STDEV({0}{1}:{0}{2})", text, this.DataStartLOC.ToString(), this.DataEndLOC.ToString());
                    break;
                case TitleTYpe.Cpu:
                    result = string.Format("({0}{1}-{0}{2})/(3*{0}{3})", new object[]
                    {
                    text,
                    this.MAX_SPC.ToString(),
                    this.AVG.ToString(),
                    this.STD.ToString()
                    });
                    break;
                case TitleTYpe.Cpl:
                    result = string.Format("({0}{1}-{0}{2})/(3*{0}{3})", new object[]
                    {
                    text,
                    this.AVG.ToString(),
                    this.MIN_SPC.ToString(),
                    this.STD.ToString()
                    });
                    break;
                case TitleTYpe.Cp_M_1:
                    result = string.Format("({0}{1}-{0}{2})/(6*{0}{3})", new object[]
                    {
                    text,
                    this.MAX_SPC.ToString(),
                    this.MIN_SPC.ToString(),
                    this.STD.ToString()
                    });
                    break;
                case TitleTYpe.Ca_L_1:
                    result = string.Format("ABS(((({0}{1}+{0}{2})/2)-{0}{3})/(({0}{1}-{0}{2})/2))", new object[]
                    {
                    text,
                    this.MAX_SPC.ToString(),
                    this.MIN_SPC.ToString(),
                    this.AVG.ToString()
                    });
                    break;
                case TitleTYpe.Cpk_M_1:
                    result = string.Format("(1-{0}{1})*{0}{2}", text, this.Ca_L_1.ToString(), this.Cp_M_1.ToString());
                    break;
                case TitleTYpe.Result:
                    result = string.Format("IF({0}{1}<1.33,\"FAIL\",\"PASS\")", text, this.Cpk_M_1);
                    break;
            }
            return result;
        }

        // Token: 0x06000003 RID: 3 RVA: 0x000023E8 File Offset: 0x000005E8
        private static string Number2ExcelCol(int source)
        {
            int num = source;
            string text = "";
            do
            {
                int num2 = num % 26;
                num /= 26;
                text = (char)(num2 + (string.IsNullOrEmpty(text) ? 65 : 64)) + text;
            }
            while (num > 26);
            if (num != 0)
            {
                text = ((char)(num + 64)).ToString() + text;
            }
            return text;
        }

        // Token: 0x0400000F RID: 15
        private int DataStartLOC;

        // Token: 0x04000010 RID: 16
        private int DataEndLOC;

        // Token: 0x04000011 RID: 17
        private int MAX_SPC = 6;

        // Token: 0x04000012 RID: 18
        private int MIN_SPC = 7;

        // Token: 0x04000013 RID: 19
        private int UNIT = 8;

        // Token: 0x04000014 RID: 20
        private int MAX = 9;

        // Token: 0x04000015 RID: 21
        private int MIN = 10;

        // Token: 0x04000016 RID: 22
        private int AVG = 11;

        // Token: 0x04000017 RID: 23
        private int STD = 12;

        // Token: 0x04000018 RID: 24
        private int Cpu = 13;

        // Token: 0x04000019 RID: 25
        private int Cpl = 14;

        // Token: 0x0400001A RID: 26
        private int Cp_M_1 = 15;

        // Token: 0x0400001B RID: 27
        private int Ca_L_1 = 16;

        // Token: 0x0400001C RID: 28
        private int Cpk_M_1 = 17;

        // Token: 0x0400001D RID: 29
        private int Result = 18;
    }
}
