using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ThomasLevesque.AutoRunCustomTool
{

    public enum UpdateTTMode
    {
        [Description("When project tree changed")]
        ProjectItemsChanged,
        [Description("When files changed")]
        FileChanged,
        [Description("When external assembly rebuilt")]
        AssemblyChanged
    }

    [CLSCompliant(false)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRebuildTTPropertyExtender
    {
        [Category("AutoRebuildTT")]
        [DisplayName("Run custom tool on")]
        [Description("When this file is saved, the custom tool will be run on the files listed in this field")]
        [Editor("System.Windows.Forms.Design.StringArrayEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        string[] RunCustomToolOn { get; set; }

        [Category("AutoRebuildTT")]
        [DisplayName("Rebuild mode")]
        [Description("Template rebuild mode")]
        [TypeConverter(typeof(EnumTypeConverter))]
        [DefaultValue(UpdateTTMode.FileChanged)]
        UpdateTTMode RunCustomToolMode { get; set; }

    }
}
