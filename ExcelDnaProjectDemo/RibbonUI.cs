using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelDnaProjectDemo
{
    /// <summary>
    /// Load custom Excel Fluent/Ribbon
    /// </summary>
    [ComVisible(true)]
    public class RibbonUI : ExcelRibbon
    {
        private static IRibbonUI customRibbon;             //记录IRibbonUI对象

        #region Fluent/Ribbon UI
        //https://blog.csdn.net/ITTechnologyHome/article/details/53891087             //VisualStudio2017集成GitHub

        //https://msdn.microsoft.com/en-us/library/aa722523(v=office.12).aspx         //Ribbon函数回调定义
        //https://msdn.microsoft.com/zh-cn/library/office/ee691833(v=office.14).aspx  //Office 2010 Backstage 视图介绍

        /// <summary>
        /// ribbon callback, get IRibbonUI object.
        /// </summary>
        public void ribbonLoaded(IRibbonUI ribbon)
        {
            customRibbon = ribbon;
        }

        /// <summary>
        /// read CustomUI.xml, xml file must be UTF-8 encode and Embedded resources.
        /// </summary>
        public override string GetCustomUI(string uiName)
        {
            string ribbonxml = string.Empty;
            try
            {
                if (ExcelDnaUtil.ExcelVersion == 12)
                    ribbonxml = ResourceHelper.GetResourceText("CustomUI12.xml");

                else
                    ribbonxml = ResourceHelper.GetResourceText("CustomUI14.xml");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return ribbonxml;
        }

        /// <summary>
        /// Ribbon callback，load image in XML element
        /// </summary>
        public override object LoadImage(string imageId)
        {
            return ResourceHelper.GetResourceBitmap(imageId);
        }

        /// <summary>
        /// ribbon callback.
        /// </summary>
        public void button_Click(IRibbonControl control)
        {
            if (control.Id.StartsWith("Menu"))
            {
                MessageBox.Show("Menu:" + control.Tag);
            }
            else
            {
                switch (control.Id)
                {                    
                    case "TestButton":
                        MessageBox.Show("Button:" + control.Id);
                        break;
                    default:
                        MessageBox.Show("Hello:" + control.Id);
                        break;
                }
            }
            //customRibbon.InvalidateControl(control.Id);
        }

        /// <summary>
        /// ribbon callback
        /// </summary>
        public stdole.IPictureDisp button_getImage(IRibbonControl control)
        {
            stdole.IPictureDisp pictureDisp = null;

            switch (control.Id)
            {
                case "TestButton":
                    pictureDisp = Image2stdoleIPictureDisp.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("unlock.png"));
                    break;
                case "TestRunTag":
                    pictureDisp = Image2stdoleIPictureDisp.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("run.png"));
                    break;
                case "ShowCTP":
                    pictureDisp = Image2stdoleIPictureDisp.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("panel.png"));
                    break;
                default:
                    pictureDisp = Image2stdoleIPictureDisp.ImageToPictureDisp(ResourceHelper.GetResourceBitmap("office.png"));
                    break;
            }
            return pictureDisp;
        }

        /// <summary>
        /// ribbon callback
        /// </summary>
        public string button_getLabel(IRibbonControl control)
        {
            string ret = string.Empty;
            switch (control.Id)
            {
                case "TestButton":
                    ret = "Test\nButton";
                    break;
                default:
                    ret = control.Id;
                    break;
            }
            return ret;
        }
        #endregion Fluent/Ribbon UI

        #region CustomTaskPane
        public void OnShowCTP(IRibbonControl control)
        {
            CTPManager.ShowCTP();
        }

        public void OnDeleteCTP(IRibbonControl control)
        {
            CTPManager.DeleteCTP();
        }
        #endregion CustomTaskPane
    }
}
