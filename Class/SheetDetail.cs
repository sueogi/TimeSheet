using System;
using System.Collections.Generic;

public class SheetList
{
    public List<SheetDetail> Sheets = new List<SheetDetail>();
}

public class SheetDetail
{
    public string SheetID { get; set; }
    public string Account { get; set; }
    public string SheetDate { get; set; }
    public string BeginHour { get; set; }
    public string BeginMin { get; set; }
    public string EndHour { get; set; }
    public string EndMin { get; set; }
    public int SpendTime { get; set; }
    public string TaskName { get; set; }
    public int TaskID { get; set; }
    public string ProjectName { get; set; }
    public string ProjectID { get; set; }
    public string StepName { get; set; }
    public int StepID { get; set; }
    public string Summary { get; set; }
    public int OverTime { get; set; }
    public DateTime? ModifyTime { get; set; }
}

