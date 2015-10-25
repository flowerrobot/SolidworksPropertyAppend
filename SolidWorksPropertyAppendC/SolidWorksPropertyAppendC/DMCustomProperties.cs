using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SwDocumentMgr;
namespace SolidWorksPropertyAppendC
{
    class DMCustomProperties
    {
        SwDMDocument18 Doc;
        List<string> _Names = null;
        ISwDMConfigurationMgr ConfigMgr;
        public DMCustomProperties(SwDMDocument18 Document)
        {
            Doc = Document;
            ConfigMgr = Doc.ConfigurationManager;
            //Document.ConfigurationManager
        }
        #region "Document"
        public List<string> GetAllDocumentPropertiesNames()
        {
            if (_Names == null || _Names.Count == 0)
            {
                _Names = new List<string>();
                if (Doc.GetCustomPropertyCount() != 0)
                {
                    foreach (string name in Doc.GetCustomPropertyNames())
                    {
                        _Names.Add(name.ToUpper());
                    }
                }
            }
            return _Names;
        }
        public bool DocumentPropertyExists(string AttributeName)
        {
            return GetAllDocumentPropertiesNames().Contains(AttributeName.ToUpper());
        }
        public SwDmCustomInfoType DocumentPropertyType(string attributeName)
        {
            SwDmCustomInfoType Type = SwDmCustomInfoType.swDmCustomInfoUnknown;
            if (DocumentPropertyExists(attributeName))
            {
                Doc.GetCustomProperty(attributeName,out Type);
                return Type;
            }
            return SwDmCustomInfoType.swDmCustomInfoUnknown;
        }
        public string DocumentPropertyEvalValue(string AttributeName)
        {
            if (DocumentPropertyExists(AttributeName))
            {
                SwDmCustomInfoType Type = SwDmCustomInfoType.swDmCustomInfoUnknown;
                string other;
                return Doc.GetCustomPropertyValues(AttributeName,out Type, out other);
            }
            return "";
        }
        public string DocumentPropertyRawValue(string AttributeName)
        {
            if (DocumentPropertyExists(AttributeName))
            {
                SwDmCustomInfoType Type = SwDmCustomInfoType.swDmCustomInfoUnknown;
                string Value = Doc.GetCustomProperty2(AttributeName, out Type);
                return Value;
            }
            return "";
        }
        public bool DocumentAddProperty(string AttributeName, SwDmCustomInfoType Type, string Value)
        {
            try
            {
                if (DocumentPropertyExists(AttributeName))
                {
                    DocumentDeleteProperty(AttributeName);
                }
                return Doc.AddCustomProperty(AttributeName, Type, Value);
            }
            catch
            {
                return false;
            }
        }
        public bool DocumentDeleteProperty(string AttributeName)
        {
            try
            {
                if (DocumentPropertyExists(AttributeName))
                {
                    return Doc.DeleteCustomProperty(AttributeName);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        public void DocumentSetProperty(string AttributeName, string Value)
        {
            Doc.SetCustomProperty(AttributeName, Value);
        }
        #endregion
        #region "Configurations"
        public List<string> GetAllPropertiesNames(SwDMConfiguration14 swCfg)
        {
            List<string> Names = new List<string>();
            foreach (string name in swCfg.GetCustomPropertyNames())
            {
                Names.Add(name.ToUpper());
            }
            foreach (string name in Doc.GetCustomPropertyNames())
            {
                if (!Names.Contains(name))
                {
                    _Names.Add(name.ToUpper());
                }
            }
            return Names;
        }
        public List<string> GetAllConfigPropertiesNames(SwDMConfiguration14 swCfg)
        {
            List<string> Names = new List<string>();
            if (swCfg.GetCustomPropertyCount() != 0)
            {
                foreach (string name in swCfg.GetCustomPropertyNames())
                {
                    Names.Add(name.ToUpper());
                }
            }
            return Names;
        }
        public bool ConfigPropertyExists(string AttributeName, SwDMConfiguration14 swCfg, bool CheckDocument = false)
        {
            bool Res = GetAllConfigPropertiesNames(swCfg).Contains(AttributeName.ToUpper());
            if (!Res && CheckDocument)
            {
                Res = this.DocumentPropertyExists(AttributeName);
            }
            return Res;
        }
        public SwDmCustomInfoType ConfigPropertyType(string attributeName, SwDMConfiguration14 swCfg)
        {
            SwDmCustomInfoType Type = default(SwDmCustomInfoType);
            //SwDmCustomInfoType = SwDmCustomInfoType.swDmCustomInfoUnknown
            if (ConfigPropertyExists(attributeName, swCfg))
            {
                swCfg.GetCustomProperty2(attributeName,out Type);
            }
            else if (this.DocumentPropertyExists(attributeName))
            {
                return this.DocumentPropertyType(attributeName);
            }
            return Type;
        }
        public string ConfigPropertyRawValue(string AttributeName, SwDMConfiguration14 swCfg)
        {
            if (ConfigPropertyExists(AttributeName, swCfg))
            {
                SwDmCustomInfoType Type = default(SwDmCustomInfoType);
                return swCfg.GetCustomProperty2(AttributeName, out Type);
                //  Return swCfg.GetCustomPropertyValues(AttributeName, Type, "")
            }
            else if (this.DocumentPropertyExists(AttributeName))
            {
                return this.DocumentPropertyEvalValue(AttributeName);
            }
            return "";
        }
        public string ConfigPropertyEvalValue(string AttributeName, SwDMConfiguration14 swCfg)
        {
            if (ConfigPropertyExists(AttributeName, swCfg))
            {
                SwDmCustomInfoType Type = default(SwDmCustomInfoType);
                string Link = "";
                return swCfg.GetCustomPropertyValues(AttributeName, out Type, out Link);
            }
            else if (this.DocumentPropertyExists(AttributeName))
            {
                return this.DocumentPropertyEvalValue(AttributeName);
            }
            return "";
        }
        public bool ConfigDeleteProperty(string AttributeName, SwDMConfiguration14 swCfg)
        {
            try
            {
                if (ConfigPropertyExists(AttributeName, swCfg))
                {
                    return swCfg.DeleteCustomProperty(AttributeName);
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        public bool ConfigAddProperty(string AttributeName, SwDmCustomInfoType Type, string Value, SwDMConfiguration14 swCfg)
        {
            ConfigDeleteProperty(AttributeName, swCfg);
            return swCfg.AddCustomProperty(AttributeName, Type, Value);
        }
        public void ConfigSetProperty(string AttributeName, string Value, SwDMConfiguration14 swCfg)
        {
            swCfg.SetCustomProperty(AttributeName, Value);
        }
        #endregion
    }
}
