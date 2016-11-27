using EnvDTE;
using Microsoft.VisualStudio;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThomasLevesque.AutoRunCustomTool
{
    public class RebuildTTExtenderProvider : IExtenderProvider
    {
        public const string SupportedExtenderName = "RebuildTTExtender";
        public const string SupportedExtenderCATID = VSConstants.CATID.CSharpFileProperties_string;
        private const string TextTemplateFileExtension = ".tt";

        public object GetExtender(string extenderCATID, string extenderName, object extendeeObject, IExtenderSite extenderSite, int cookie)
        {
            IRebuildTTPropertyExtender extender = null;

            if (this.CanExtend(extenderCATID, extenderName, extendeeObject))
            {
                var fileName = extendeeObject.GetPropertyValue<string>("FileName");
                extender = new RebuildTTPropertyExtender(fileName, extenderSite, cookie);
            }

            return extender;
        }

        public bool CanExtend(string extenderCATID, string extenderName, object extendeeObject)
        {
            return extenderName == SupportedExtenderName // check if the correct extener is requested
                && string.Equals(extenderCATID, SupportedExtenderCATID, StringComparison.OrdinalIgnoreCase) // check if the correct CATID is requested
                && string.Equals(extendeeObject.GetPropertyValue<string>("ExtenderCATID"), SupportedExtenderCATID, StringComparison.OrdinalIgnoreCase) // check whether the extended object really has this CATID
                && string.Equals(extendeeObject.GetPropertyValue<string>("Extension"), TextTemplateFileExtension, StringComparison.OrdinalIgnoreCase); // check whether the file extension is '.tt'
        }
    }

    public static class TypeDescriptorSupport
    {
        public static T GetPropertyValue<T>(this object source, string propertyName)
        {
            if (source == null)
            {
                throw new ArgumentNullException("source");
            }

            if (propertyName == null)
            {
                throw new ArgumentNullException("propertyName");
            }

            object value = null;

            System.ComponentModel.PropertyDescriptor property = System.ComponentModel.TypeDescriptor.GetProperties(source)[propertyName];
            if (property != null)
            {
                value = property.GetValue(source);
            }

            return value != null ? (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture) : default(T);
        }
    }
}
