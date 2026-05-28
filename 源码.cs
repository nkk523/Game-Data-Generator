// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using System.Drawing;

ExcelPackage.License.SetNonCommercialPersonal("游戏道具配置表自动生成与智能校验工具.xlsx");

using (var package=new ExcelPackage())
{ 
    var worksheet=package.Workbook.Worksheets.Add("游戏道具配置表");


    //填充表头
    worksheet.Cells["A1"].Value= "道具ID";
    worksheet.Cells["B1"].Value = "道具名称";
    worksheet.Cells["C1"].Value= "道具类型";
    worksheet.Cells["D1"].Value = "价格";
    worksheet.Cells["E1"].Value = "稀有度";


    //功能一：防编号重复
    var range = worksheet.Cells["A2:A1000"];

    var rule = range.ConditionalFormatting.AddExpression();          //表达式制定规则
    rule .Formula=$"COUNTIF($A$2:$A$1000,A2)>1";                     //满足出现次数大于一
    rule.Style.Fill.BackgroundColor.Color=Color.Red;                 //满足规则则标红背景


    //功能二：下拉菜单
    var validation = worksheet.DataValidations.AddListValidation("C2:C1000");  //指定菜单列
    validation.Formula.Values.Add("消耗品");                                   //三个选项
    validation.Formula.Values.Add("装备"); 
    validation.Formula.Values.Add("材料");

    //功能三：下拉菜单＋稀有度标记区分
    var validation1 = worksheet.DataValidations.AddListValidation("E2:E1000");
    validation1.Formula.Values.Add("五星");
    validation1.Formula.Values.Add("四星");
    validation1.Formula.Values.Add("三星");


    var range1=worksheet.Cells["E2:E1000"];
    var rule1=range1.ConditionalFormatting.AddExpression();         //五星稀有度
    rule1.Formula = "E2=\"五星\"";
    rule1.Style.Fill.BackgroundColor.Color=Color.Red;
    rule1.Style.Font.Color.Color=Color.Gold;

    var rule2=range1.ConditionalFormatting.AddExpression();         //四星稀有度
    rule2.Formula = "E2=\"四星\"";
    rule2 .Style.Fill.BackgroundColor.Color = Color.Pink;
    rule2.Style.Font.Color.Color=Color.Black;

    var rule3=range1.ConditionalFormatting.AddExpression();         //三星稀有度
    rule3.Formula = "E2=\"三星\"";
    rule3.Style.Fill.BackgroundColor.Color=Color.Blue;
    rule3 .Style.Font.Color.Color = Color.White;


        //保存excel文件
    var fileInfo = new FileInfo(@"E:\项目\练习\游戏道具配置表自动生成与智能校验工具.xlsx");
    using (var filestream=fileInfo.OpenWrite())
    { 
        package.SaveAs(filestream);
    }

}