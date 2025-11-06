
var Cmd;
//-----------------------------------------------------------
function inpath(file)
{
    var ForReading = 1, ForWriting = 2;
    var wsh = new ActiveXObject("wscript.shell");
    var env=wsh.Environment("SYSTEM");
    var path=env.item("SAKURA_SCRIPT") + "/";

    var FileOpener = new ActiveXObject( "Scripting.FileSystemObject");
    var FilePointer = FileOpener.OpenTextFile(path + file, ForReading, true);
    Cmd = FilePointer.ReadAll();
}

//-----------------------------------------------------------
inpath("inc.js");
eval(Cmd);

var fso=new ActiveXObject("Scripting.FileSystemObject");

var show_txt="";
make_XL();

// chart_1();
// chart_2();
get_chart_property()

function chart_1(){
    var wb=XL.ActiveWorkbook;
    var ws=XL.ActiveSheet;

    // 写入一些数据
    ws.Cells(1, 1).Value = "月份";
    ws.Cells(1, 2).Value = "销售额";
    ws.Cells(2, 1).Value = "一月";
    ws.Cells(2, 2).Value = 120;
    ws.Cells(3, 1).Value = "二月";
    ws.Cells(3, 2).Value = 150;
    ws.Cells(4, 1).Value = "三月";
    ws.Cells(4, 2).Value = 180;

    /*
    var chart = wb.Charts.Add();
    chart.ChartType = 51; // 柱状图
    chart.SetSourceData(ws.Range("A1:B4"));
    chart.HasTitle = true;
    chart.ChartTitle.Text = "季度销售统计";
    */

    var dataRange = ws.Range("A1:B4");

    var chartObj = ws.ChartObjects().Add(200, 50, 400, 300); 
    // 参数：(Left, Top, Width, Height)

    // 获取嵌入的 Chart 对象
    var chart = chartObj.Chart;

    // 绑定数据源
    chart.SetSourceData(dataRange);

    // 设置图表类型（柱状图）
    chart.ChartType = 51; // xlColumnClustered = 51

    // 设置标题
    chart.HasTitle = true;
    chart.ChartTitle.Text = "季度销售图";
}
function chart_2(){
    var wb=XL.ActiveWorkbook;
    var ws=XL.ActiveSheet;

    // 写入测试数据
    ws.Cells(1, 1).Value = "月份";
    ws.Cells(1, 2).Value = "销售额A";
    ws.Cells(1, 3).Value = "销售额B";
    ws.Cells(2, 1).Value = "一月";
    ws.Cells(3, 1).Value = "二月";
    ws.Cells(4, 1).Value = "三月";
    ws.Cells(2, 2).Value = 120;
    ws.Cells(3, 2).Value = 150;
    ws.Cells(4, 2).Value = 180;
    ws.Cells(2, 3).Value = 200;
    ws.Cells(3, 3).Value = 250;
    ws.Cells(4, 3).Value = 300;

    // 在工作表上插入一个图表对象
    var chartObj = ws.ChartObjects().Add(200, 50, 450, 300);
    var chart = chartObj.Chart;

    // 设置图表类型（折线图）
    // chart.ChartType = 4; // xlLine
    chart.ChartType = 51; // xlColumnClustered = 51

    // 清空默认系列
    while (chart.SeriesCollection().Count > 0) {
        chart.SeriesCollection(1).Delete();
    }

    // ✅ 手动添加数据系列1
    var series1 = chart.SeriesCollection().NewSeries();
    series1.Name = "=\"销售额A\"";
    series1.XValues = ws.Range("A2:A4");  // 横轴
    series1.Values = ws.Range("B2:B4");   // 纵轴

    // ✅ 手动添加数据系列2
    var series2 = chart.SeriesCollection().NewSeries();
    series2.Name = "=\"销售额B\"";
    series2.XValues = ws.Range("A2:A4");
    series2.Values = ws.Range("C2:C4");

    // 设置标题
    chart.HasTitle = true;
    chart.ChartTitle.Text = "各月销售趋势图";

    // 设置轴标题（可选）
    chart.Axes(1).HasTitle = true;  // 横轴
    chart.Axes(1).AxisTitle.Text = "月份";
    chart.Axes(2).HasTitle = true;  // 纵轴
    chart.Axes(2).AxisTitle.Text = "销售额";
}
function get_chart_property(){

    var wb=XL.ActiveWorkbook;
    var ws=XL.ActiveSheet;

    var charts = ws.ChartObjects();
    var count = charts.Count;

    if (count === 0) {
        WScript.Echo("此工作表中没有图表对象。");
    } else {
        WScript.Echo("图表总数: " + count);
        for (var i = 1; i <= count; i++) {
            var obj = charts.Item(i);
            var chart = obj.Chart;

            WScript.Echo("---------------------------");
            WScript.Echo("索引: " + i);
            WScript.Echo("名称: " + obj.Name);
            WScript.Echo("Left: " + obj.Left + ", Top: " + obj.Top);
            WScript.Echo("宽度: " + obj.Width + ", 高度: " + obj.Height);
            WScript.Echo("ChartType: " + chart.ChartType);
            WScript.Echo("HasTitle: " + chart.HasTitle);
            if (chart.HasTitle)
                WScript.Echo("Title: " + chart.ChartTitle.Text);
        }
    }
}

