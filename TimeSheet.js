var Tasks = [{ OptionText: "請選擇", OptionValue: "0" }];
var Projects = [{ OptionText: "無", OptionValue: "0" }];
var Steps = [{ OptionText: "無", OptionValue: "0" }];
var Options = [Tasks, Projects, Steps];
var ViewModel;
var mySheetData;
var chosenDate;
var OnTime;
var OffTime;

$(function () {
    var EditableDays;
    var PageLimit;
    var today = new Date();

    // slice(-2) 的意思是取此值從後往前算兩位，類似 substring(str.length-2, str,length-1)
    // 用這個方法可以讓非兩位數的月份前面補零，但是已經是兩位數的月份不加零
    var thisday = today.getFullYear() + "/" + ("0" + today.getMonth() + 1).slice(-2) + "/" + ("0" + today.getDate()).slice(-2);

    // collapse top header
    $("#header_image").slideUp(1000);
    $("#menu-wrapper").animate({ top: "0px" }, 1000);

    // today
    $("#today").text((today.getMonth() + 1) + "/" + ("0" + today.getDate()).slice(-2));

    // get system parameter
    $.ajax({
        type: "POST",
        url: "GetInitialHandler.ashx",
        dataType: "json",
        // 要先拿到系統參數才能做後續，所以把非同步關掉
        async: false,
        success: function (data) {
            EditableDays = data.EditableDays;
            PageLimit = data.PageLimit;
            OnTime = data.OnTime;
            OffTime = data.OffTime;
        }
    }).fail(function () { alert("系統異常，請通知系統管理員。#1"); return false; });

    // 取一周摘要
    GetWeeks(thisday);

    // 這個class用來當目前頁面的flag
    $("#week").show().addClass("page_open");
    $("#sheet").hide();

    // make footer
    for (var i = EditableDays - 1; i >= 0; i--) {
        // 做出頁籤物件，這樣$("<li>")宣告會自動做出開始結束標籤 <li></li>
        var tab = $("<li>");
        // <li></li>裡面的文字
        // 利用datepicker，格式是mm/dd，因為今天要放在最後面，所以用-i
        tab.append($.datepicker.formatDate("mm/dd", new Date(new Date().setDate(new Date().getDate() - i))));
        // 這邊設id是為了如果這幾個頁籤有跨年，查詢條件資料從id取就不會有年不對的問題
        tab.attr("id", $.datepicker.formatDate("yy/mm/dd", new Date(new Date().setDate(new Date().getDate() - i))));
        tab.addClass("navbar_tab");
        $("#footer_navbar").find("ul").append(tab);
    }
    // 最後一個頁籤加上css
    $("#footer_navbar").find("ul").find("li:last-child").addClass("today_tab");

    // footer click
    $("#footer_navbar").delegate("li", "click", function () {
        $("#today").text($(this).text());

        // 選擇的頁籤加上與其他頁籤不一樣的css
        $(this).addClass("selected_tab");
        // 如果換別的頁籤，簡單的方法就是這個以外的頁籤都移除這個css
        $("#footer_navbar").find("li").not(this).removeClass("selected_tab");

        if ($(this).text() == "一周摘要") {
            // 因為資料可能會異動，所以重新取一次資料
            GetWeeks(thisday);
            $("#today").hide();
            $("#sheet").hide("slide", { direction: "left" }, 1000).removeClass("page_open");
            $("#week").show("slide", { direction: "left" }, 1000).addClass("page_open");
        } else {
            $("#today").show();
            // 如果上一個開的頁籤是一周摘要，就把flag移除。                   
            if ($("#week").hasClass("page_open")) {
                $("#week").hide("slide", { direction: "left" }, 1000).removeClass("page_open");
                $("#sheet").show("slide", { direction: "left" }, 1000).addClass("page_open");
            } else {
                $("#sheet").show("slide", { direction: "left" }, 1000).addClass("page_open");
            }

            chosenDate = $(this).attr("id");

            // 如果已經開過某天的頁籤，代表ViewModel已經建立過，就清空再抓新資料
            if (typeof (ViewModel) != "undefined") {
                mySheetData.removeAll();
                getSheetData(chosenDate, PageLimit);
            } else {
                // 沒有建立過ViewModel就new一個出來，並且綁定資料
                ViewModel = new SheetViewModel();
                getSheetData(chosenDate, PageLimit);
                ko.applyBindings(ViewModel);
            }
        }
    }); // end of $("#footer_navbar") delegate

    // when focus row id change then save this row            
    var currentRow;
    var previousRow;
    $("#sheet_tbody").delegate("tr", {
        focus: function () {
            // 如果focus的這一列的id跟currentRow不一樣，而且currentRow有過記錄，就把上一列的資料丟去save
            if ($(this).attr("id") != currentRow && typeof (currentRow) != "undefined") {
                // 如果用jquery方法去取會變成jquery物件，不能直接轉為knockout物件
                // 但用原生的getElementById就沒問題
                var context = ko.contextFor(document.getElementById(previousRow));
                saveSheet(context);
            }
            currentRow = $(this).attr("id");
            // ko.contextFor這個動作可以取到knockout物件而不是javascript物件
            var context = ko.contextFor(this);
            $("#row_img" + context.$index()).removeAttr("title");
            context.$data.imgpath("../Image/writing.png");

            context.$data.editing(true);
        },
        blur: function () {
            previousRow = $(this).attr("id");

            var context = ko.contextFor(this);
            context.$data.editing(false);
        }
    });

    // get project, task, step
    $.ajax({
        type: "POST",
        url: "GetSelectHandler.ashx",
        dataType: "json",
        async: false,
        success: function (data) {
            var count = 0;
            $.each(data.Items, function (index, value) {
                // 取出data.Items裡面的所有最高層物件的名字，目前的狀況就是拿到 "Task", "Proj", "Step"
                // 這樣寫的好處是以後可以再多加不同類的下拉選單而不用改這一段的程式
                for (var i in data.Items) {
                    // 不知道原因程式會跑四次，所以只要count到3就強制結束
                    // 試過只有Object.keys(data.Items).length才能取到data.Items的length
                    if (count == Object.keys(data.Items).length) break;
                    $.each(data.Items[i], function (index, value) {
                        // 把data.Items.XXX的內容物依序push到Options
                        Options[count].push({ OptionText: this.OptionName, OptionValue: this.OptionId });
                    });
                    count++;
                }
            });
        }
    }).fail(function () { alert("系統異常，請通知系統管理員。#2"); return false; });
}) // end of onload

function GetWeeks(thisday) {
    $("#week_tbody").children().remove();
    $("#everyday_tbody").children().remove();

    $.getJSON("GetWeekHandler.ashx", {}, function (data) {
        $.each(data, function (index, value) {
            var $tr = $("<tr>");
            $tr.append([$("<td>").text(value.KeyName), $("<td>").text(value.NormalTime).addClass("text_right"), $("<td>").text(value.OverTime).addClass("text_right")]);
            $("#week_tbody").append($tr);
        });
    });

    $.getJSON("GetWeekHandler.ashx", { date: thisday }, function (data) {
        $.each(data, function (index, value) {
            var $tr = $("<tr>");
            $tr.append([$("<td>").text(value.KeyName), $("<td>").text(value.NormalTime).addClass("text_right"), $("<td>").text(value.OverTime).addClass("text_right")]);
            $("#everyday_tbody").append($tr);
        });
    });
}

function getSheetData(date, rowcount) {
    $.getJSON("GetSheetHandler.ashx", { date: date }, function (data) {
        console.log("1");
        $.each(data.Sheets, function (index, value) {
            // 依name來找到下拉選單是選第幾個
            // 這裡不能用迴圈做的關鍵在於value.xxxName，如果這個xxxName可以array化就可以用迴圈做
            var task_index, proj_index, step_index;
            getIndexByKey(Options[0], value.TaskName, function (_index) { task_index = _index; });
            getIndexByKey(Options[1], value.ProjectName, function (_index) { proj_index = _index; });
            getIndexByKey(Options[2], value.StepName, function (_index) { step_index = _index; });
            mySheetData.push(new Sheet(this.SheetID, this.SheetDate, this.BeginHour, this.BeginMin, this.EndHour, this.EndMin, this.SpendTime, this.OverTime, ViewModel.availableTasks[task_index], this.TaskID, ViewModel.availableProjects[proj_index], this.ProjectID, ViewModel.availableSteps[step_index], this.StepID, this.Summary));
        });
    }).done(function () {
        // 計算資料筆數，不足筆數用空白填充
        for (var i = mySheetData().length; i < rowcount; i++) {
            mySheetData.push(new Sheet("", "", "", "", "", "", 0, 0, ViewModel.availableTasks[0], "", ViewModel.availableProjects[0], "", ViewModel.availableSteps[0], "", ""));
        }
        $("#sheet_tbody").find("tr:odd").addClass("tr_odd").end().find("tr:even").addClass("tr_even");
    }).fail(function () { alert("系統異常，請通知系統管理員。#3"); return false; });
}

// knockout的ViewModel
function SheetViewModel() {
    var self = this;
    self.availableTasks = Options[0];
    self.availableProjects = Options[1];
    self.availableSteps = Options[2];
    mySheetData = ko.observableArray();
    self.sheets = mySheetData;

    // 新增一列空白的TimeSheet，這裡的css還有問題
    self.addSheet = function () {
        mySheetData.push(new Sheet("", "", "", "", "", "", 0, 0, ViewModel.availableTasks[0], "", ViewModel.availableProjects[0], "", ViewModel.availableSteps[0], "", ""));
    };

    // 用ko.computed這個方法綁定資料才會即時update
    self.totalSpendTime = ko.computed(function () {
        var total = 0;
        $.each(self.sheets(), function (index, value) {
            total += value.spendtime();
        });
        return total;
    });

    self.totalOverTime = ko.computed(function () {
        var total = 0;
        $.each(self.sheets(), function (index, value) {
            total += value.overtime();
        });
        return total;
    });

} // end of SheetViewModel

function Sheet(SheetID, SheetDate, BeginHour, BeginMin, EndHour, EndMin, SpendTime, OverTime, Task, TaskID, Project, ProjectID, Step, StepID, Summary) {
    var self = this;
    // ko.observable也是綁定資料即時update
    self.sheetid = ko.observable(SheetID);
    self.sheetdate = ko.observable(SheetDate);

    // extend是擴充方法，用來驗證輸入
    // 0跟1是分辨hour or minute
    self.beginhour = ko.observable(BeginHour).extend({ numeric: 0 });
    self.beginmin = ko.observable(BeginMin).extend({ numeric: 1 });
    self.endhour = ko.observable(EndHour).extend({ numeric: 0 });
    self.endmin = ko.observable(EndMin).extend({ numeric: 1 });

    // 計算工作費時
    self.spendtime = ko.computed(function () {
        var spend = ((self.endhour() * 60) + (parseInt(self.endmin()))) -
            ((self.beginhour() * 60) + (parseInt(self.beginmin())));
        return isNaN(spend) || spend <= 0 ? 0 : spend;
    });

    self.overtime = ko.computed(function () {
        // 把上下班時間的小時跟分鐘分開來比較好比較
        var ot = parseTime(OnTime);
        var oft = parseTime(OffTime);
        var over;
        // 加班與否的邏輯
        if (self.beginhour() >= oft.hh) {
            over = (new Date(2000, 1, 1, self.endhour(), self.endmin()) - new Date(2000, 1, 1, self.beginhour(), self.beginmin())) / (1000 * 60);
        }
        if (self.endhour() >= oft.hh && self.endmin() >= oft.mm && self.beginhour() < oft.hh) {
            over = (new Date(2000, 1, 1, self.endhour(), self.endmin()) - new Date(2000, 1, 1, oft.hh, oft.mm)) / (1000 * 60);
        }
        if (self.beginhour() < ot.hh && self.beginhour() != "") {
            over = (new Date(2000, 1, 1, ot.hh, ot.mm) - new Date(2000, 1, 1, self.beginhour(), self.beginmin())) / (1000 * 60);
        }
        if (self.beginhour() < ot.hh && self.endhour() < ot.hh) {
            over = (new Date(2000, 1, 1, self.endhour(), self.endmin()) - new Date(2000, 1, 1, self.beginhour(), self.beginmin())) / (1000 * 60);
        }
        return isNaN(over) || over < 0 ? 0 : over;
    });

    self.task = ko.observable(Task);
    self.taskid = ko.computed(function () { return self.task().OptionValue; });

    self.project = ko.observable(Project);
    self.projectid = ko.computed(function () { return self.project().OptionValue; });

    self.step = ko.observable(Step);
    self.stepid = ko.computed(function () { return self.step().OptionValue; });

    self.summary = ko.observable(Summary);

    self.imgpath = ko.observable("../Image/sheet.png");

    self.editing = ko.observable(false);
} // end of Sheet

// 驗證時間輸入，看knockout官網範例會比較清楚這裡在做啥
ko.extenders.numeric = function (target, precise) {
    var result = ko.computed({
        read: target,  //always return the original observables value
        write: function (newValue) {
            var current = target(),
                newValueAsHour = (isNaN(newValue) || newValue < 0 || newValue > 23) ? "" : newValue,
                newValueAsMinute = (isNaN(newValue) || newValue < 0 || newValue > 59) ? "" : newValue,
                valueToWrite = precise == 0 ? newValueAsHour : newValueAsMinute;
            //only write if it changed
            if (valueToWrite !== current) {
                target(valueToWrite);
            } else {
                //if the rounded value is the same, but a different value was written, force a notification for the current field
                if (newValue !== current) {
                    target.notifySubscribers(valueToWrite);
                }
            }
        }
    }).extend({ notify: 'always' });
    //initialize with current value to make sure it is rounded appropriately
    result(target());
    //return the new computed observable
    return result;
};

function saveSheet(context) {
    var self = context.$data;

    // 儲存與否的邏輯，這裡可能不夠周全
    if (self.beginhour() && self.beginmin() && self.endhour() && self.endmin() &&
        //(self.beginhour() == self.endhour()) && self.beginmin() < self.endmin() && 
        self.task().OptionValue != "0") {

        if (self.sheetdate() == "") {
            self.sheetdate(chosenDate);
        }

        var sheet = ko.toJSON({
            // Account因為在javascript裡取有"\"的session會有問題，所以留到後端去取session，這裡隨便給值
            "Account": "workway",
            "SheetID": self.sheetid(),
            "SheetDate": self.sheetdate(),
            "BeginHour": self.beginhour(),
            "BeginMin": self.beginmin(),
            "EndHour": self.endhour(),
            "EndMin": self.endmin(),
            "TaskID": self.taskid(),
            "ProjectID": self.projectid(),
            "StepID": self.stepid(),
            "Summary": self.summary()
        });

        $.ajax({
            type: "POST",
            url: "SaveSheetHandler.ashx",
            data: sheet,
            dataType: "text",
            async: false,
            success: function (result) {
                if (result != "false") {
                    self.SheetID = result;
                    self.imgpath("../Image/write_success.png");
                } else {
                    self.imgpath("../Image/write_failure.png");
                }
            },
            error: function () {
                alert("系統異常，請通知系統管理員。#4");
            }
        });
    } else {
        $("#row_img" + context.$index()).attr("title", "必填欄位未完成");
        self.imgpath("../Image/write_failure.png");
    }
}

// function tool
function parseTime(s) {
    var part = s.match(/(\d+):(\d+)(?: )?/i);
    var hh = parseInt(part[1], 10);
    var mm = parseInt(part[2], 10);
    var ap = part[3] ? part[3].toUpperCase() : null;
    return { hh: hh, mm: mm };
}

function getIndexByKey(obj, key, callback) {
    $.each(obj, function (index, value) {
        // 因為直接return會有問題，但是回傳一個方法就ok
        // 也許可以改成一找到相對的值就break
        if (key == value.OptionText) { callback(index); }
    });
}
