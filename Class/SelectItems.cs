using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// SelectItems 的摘要描述
/// </summary>
public class SelectItems
{
    public ItemCollection Items = new ItemCollection();
}

public class ItemCollection
{
    public List<Option> Task = new List<Option>();
    public List<Option> Proj = new List<Option>();
    public List<Option> Step = new List<Option>();    
}

public class Option
{
    public string OptionId { get; set; }
    public string OptionName { get; set; }
    public int OptionSort { get; set; }
}