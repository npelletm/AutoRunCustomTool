using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ThomasLevesque.AutoRunCustomTool
{
    [CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class RebuildTTPropertyExtender : IDisposable, IRebuildTTPropertyExtender
    {
        private readonly ProjectItem _projectItem;
        private readonly IExtenderSite _extenderSite;
        private readonly int _cookie;

        private bool _disposed = false;

        #region Ctor, dtor
        public RebuildTTPropertyExtender(string fileName, IExtenderSite extenderSite, int cookie)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException("fileName");
            }

            if (extenderSite == null)
            {
                throw new ArgumentNullException("extenderSite");
            }

            this._extenderSite = extenderSite;
            this._cookie = cookie;

            // resolve the project item from the file name
            var dte = (DTE)Package.GetGlobalService(typeof(DTE));
            this._projectItem = dte.Solution.FindProjectItem(fileName);
        }

        ~RebuildTTPropertyExtender()
        {
            this.Dispose(false);
        }
        #endregion


        //[DefaultValue(false)]
        //[DisplayName("Exclude StyleCop")]
        //[Category("StyleCop")]
        //[Description("Specifies that the file is exculded from the StyleCop source analysis.")]
        //public bool ExcludeStyleCop
        //{
        //    get
        //    {
        //        return this._projectItem.GetItemAttribute<string>("ExcludeStyleCop");
        //    }

        //    set
        //    {
        //        this._projectItem.SetItemAttribute("ExcludeStyleCop", value);
        //    }
        //}

        [Category("AutoRebuildTT")]
        [DisplayName("Run custom tool on")]
        [Description("When this file is saved, the custom tool will be run on the files listed in this field")]
        [Editor("System.Windows.Forms.Design.StringArrayEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]

        public string[] RunCustomToolOn
        {
            get
            {
                var s = this._projectItem.GetItemAttribute<string>(AutoRunCustomToolPackage.TargetsPropertyName);
                if (s != null)
                {
                    return s.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                }
                return null;
            }
            set
            {
                string s = null;
                if (value != null)
                {
                    s = string.Join(";", value);
                }
                this._projectItem.SetItemAttribute(AutoRunCustomToolPackage.TargetsPropertyName, s);
            }
        }

        //pour un enuméré

        [Category("AutoRebuildTT")]
        [DisplayName("Rebuild mode")]
        [Description("Template rebuild mode")]
        [TypeConverter(typeof(EnumTypeConverter))]
        [DefaultValue(UpdateTTMode.FileChanged)]
        public UpdateTTMode RunCustomToolMode
        {
            get
            {
                var s = this._projectItem.GetItemAttribute<string>("RunCustomToolMode");
                if (s != null)
                {
                    UpdateTTMode mode;
                    if (Enum.TryParse<UpdateTTMode>(s, out mode))
                        return mode;
                }
                return UpdateTTMode.FileChanged;
            }
            set
            {
                string s = null;
                if (Enum.IsDefined(typeof(UpdateTTMode), value))
                {
                    s = value.ToString();
                }
                this._projectItem.SetItemAttribute("RunCustomToolMode", s);
            }
        }

        #region IDisposable
        public void Dispose()
        {
            this.Dispose(true);

            // take the instance off of the finalization queue.
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // check to see if Dispose has already been called.
            if (!this._disposed)
            {
                // if IDisposable.Dispose was called, dispose all managed resources
                if (disposing)
                {
                    // free other state (Managed objcts)
                    if (this._cookie != 0)
                    {
                        this._extenderSite.NotifyDelete(this._cookie);
                    }
                }

                // if Finilize or IDisposable.Dispose, free your own state (unmanaged objects)
                this._disposed = true;
            }
        }
        #endregion
    }


    public static class ProjectItemSupport
    {
        public static void SetItemAttribute<T>(this ProjectItem item, string attributeName, T value)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }

            if (attributeName == null)
            {
                throw new ArgumentNullException("attributeName");
            }

            IVsHierarchy hierarchy = item.GetProjectOfUniqueName();
            var buildPropertyStorage = hierarchy as IVsBuildPropertyStorage;
            if (buildPropertyStorage != null)
            {
                uint itemId = hierarchy.ParseCanonicalName(item);
                var attributeValue = (string)Convert.ChangeType(value, typeof(string), CultureInfo.InvariantCulture);

                int hresult = buildPropertyStorage.SetItemAttribute(itemId, attributeName, attributeValue);
                if (hresult != 0)
                {
                    string message = string.Format(CultureInfo.InvariantCulture, "Could not set the value {0} for item attribute {1}", attributeValue, attributeName);
                    throw new InvalidOperationException(message);
                }
            }
        }

        public static T GetItemAttribute<T>(this ProjectItem item, string attributeName)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }

            if (attributeName == null)
            {
                throw new ArgumentNullException("attributeName");
            }

            string value = string.Empty;
            IVsHierarchy hierarchy = item.GetProjectOfUniqueName();

            var buildPropertyStorage = hierarchy as IVsBuildPropertyStorage;
            if (buildPropertyStorage != null)
            {
                uint itemId = hierarchy.ParseCanonicalName(item);

                int hresult = buildPropertyStorage.GetItemAttribute(itemId, attributeName, out value);
                if (hresult != 0)
                {
                    //string message = string.Format(CultureInfo.InvariantCulture, "Could not get the value from the item attribute {0}", attributeName);
                    //throw new InvalidOperationException(message);
                    return default(T);
                }
            }

            return string.IsNullOrWhiteSpace(value) ? default(T) : (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
        }

        private static IVsHierarchy GetProjectOfUniqueName(this ProjectItem item)
        {
            IVsHierarchy hierarchy;
            var solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));
            int hresult = solution.GetProjectOfUniqueName(item.ContainingProject.UniqueName, out hierarchy);
            if (hresult != 0)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "Could not retrieve project hierarchy with name {0}", item.ContainingProject.UniqueName);
                throw new InvalidOperationException(message);
            }

            return hierarchy;
        }

        private static uint ParseCanonicalName(this IVsHierarchy hierarchy, ProjectItem item)
        {
            uint itemId;
            var fullPath = (string)item.Properties.Item("FullPath").Value;
            int hresult = hierarchy.ParseCanonicalName(fullPath, out itemId);
            if (hresult != 0)
            {
                string message = string.Format(CultureInfo.InvariantCulture, "Could not parse canonical name {0}", fullPath);
                throw new InvalidOperationException(message);
            }

            return itemId;
        }
    }

    public class EnumTypeConverter : EnumConverter
    {
        private Type m_EnumType;

        public EnumTypeConverter(Type type)
            : base(type)
        {
            m_EnumType = type;
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destType)
        {
            return destType == typeof(string);
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destType)
        {
            FieldInfo fi = m_EnumType.GetField(Enum.GetName(m_EnumType, value));
            DescriptionAttribute dna = (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));

            if (dna != null)
                return dna.Description;
            else
                return value.ToString();
        }

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type srcType)
        {
            return srcType == typeof(string);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            foreach (FieldInfo fi in m_EnumType.GetFields())
            {
                DescriptionAttribute dna = (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));

                if ((dna != null) && ((string)value == dna.Description))
                    return Enum.Parse(m_EnumType, fi.Name);
            }
            return Enum.Parse(m_EnumType, (string)value);
        }
    }
    //public class UpdateTTModeTypeConverter : TypeConverter
    //{
    //    public static readonly string[] Modes = Enum.GetNames(typeof(UpdateTTMode));

    //    public override bool GetStandardValuesExclusive(ITypeDescriptorContext context)
    //    {
    //        return true;
    //    }

    //    public override bool GetStandardValuesSupported(ITypeDescriptorContext context)
    //    {
    //        return true;
    //    }

    //    public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext context)
    //    {
    //        return new StandardValuesCollection(Modes);
    //    }

    //    //String -> UpdateTTMode
    //    public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
    //    {
    //        if (sourceType == typeof(string))
    //        {
    //            return true;
    //        }
    //        return base.CanConvertFrom(context, sourceType);
    //    }
    //    public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
    //    {
    //        if (value is string)
    //        {
    //            UpdateTTMode val;

    //            var ret = Enum.TryParse<UpdateTTMode>((string)value, out val);
    //            if (ret)
    //                return val;
    //        }
    //        return base.ConvertFrom(context, culture, value);
    //    }
    //    //UpdateTTMode -> String : default
    //    public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
    //    {
    //        if (destinationType == typeof(string))
    //        {
    //            return true;
    //        }
    //        return base.CanConvertTo(context, destinationType);
    //    }
    //    public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
    //    {
    //        if (destinationType == typeof(string))
    //        {
    //            return value.ToString();
    //        }
    //        return base.ConvertTo(context, culture, value, destinationType);
    //    }
    //}


    //public class UpdateTTModeUITypeEditor : System.Drawing.Design.UITypeEditor
    //{
    //    private readonly ListBox _listBox = new ListBox();
    //    private IWindowsFormsEditorService _windowsFormsEditorService;

    //    public UpdateTTModeUITypeEditor()
    //    {
    //        _listBox = new ListBox();
    //        _listBox.BorderStyle = BorderStyle.None;
    //        _listBox.Items.Clear();
    //        _listBox.Items.AddRange(UpdateTTModeTypeConverter.Modes);
    //        _listBox.SelectionMode = SelectionMode.One;

    //        //add event handler for drop-down box when item will be selected
    //        _listBox.Click += new EventHandler(Box1_Click);
    //    }

    //    public override System.Drawing.Design.UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
    //    {
    //        return UITypeEditorEditStyle.DropDown;
    //    }

    //    // Displays the UI for value selection.
    //    public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
    //    {
    //        // drop-down UI in the Properties window.
    //        this._windowsFormsEditorService = (IWindowsFormsEditorService)provider.GetService(typeof(IWindowsFormsEditorService));
    //        if (this._windowsFormsEditorService != null)
    //        {
    //            this._windowsFormsEditorService.DropDownControl(this._listBox);
    //            return this._listBox.SelectedItem;
    //        }
    //        return value;
    //    }

    //    private void Box1_Click(object sender, EventArgs e)
    //    {
    //        this._windowsFormsEditorService.CloseDropDown();
    //    }
    //}

}
