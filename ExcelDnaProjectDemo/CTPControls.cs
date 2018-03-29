using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDnaProjectDemo
{
    /// <summary>
    /// 任务窗格使用的自定义窗体控件
    /// </summary>
    [ComVisible(true)]
    public class CTPControls : UserControl
    {
        public Label TheLabel;
        public CTPControls()
        {
            TheLabel = new Label();
            TheLabel.Text = "My First CTP!";
            TheLabel.Location = new System.Drawing.Point(20, 20);
            TheLabel.Size = new System.Drawing.Size(200, 60);

            Controls.Add(TheLabel);
        }
    }
}
