using ExcelDna.Integration.CustomUI;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Excel4Engineers
{
    [ComVisible(true)]
    public abstract class RibbonControllerBase : ExcelRibbon
    {
        public abstract string TabName { get; }

        public IRibbonUI RibbonUI { get; set; }

        public override string GetCustomUI(string RibbonID)
        {
            var xml =  GenerateXml();

            return xml;
        }

        /// <summary>
        /// A method called by excel with the ribbon is loaded
        /// </summary>
        /// <param name="ribbon"></param>
        public void OnLoad(IRibbonUI ribbon)
        {
            //Ensure the IRibbonUI is cached
            RibbonUI = ribbon;

            //Activate the ribbon when it is loaded (not required)
            //ribbon.ActivateTab(TabName);
        }

        /// <summary>
        /// The current group index used to generate names of the group Ids, 
        /// ensures that no group has the same Id
        /// </summary>
        private int _GroupIndex = 0;

        /// <summary>
        /// The current group index used to generate names of the ribbon element Ids, 
        /// ensures that no UI element has the same Id
        /// </summary>
        private int _UIElementIndex = 0;

        /// <summary>
        /// Gets the next group index used to generate names of the group Ids, 
        /// Called when UI elements InitAndGetXml()
        /// </summary>
        protected int GetNextGroupIndex()
        {
            _GroupIndex++;
            return _GroupIndex;
        }

        /// <summary>
        /// Gets the next ui element index used to generate names of the ui element Ids, 
        /// Called when UI elements InitAndGetXml()
        /// </summary>
        internal int GetNextUIElementIndex()
        {
            _UIElementIndex++;
            return _UIElementIndex;
        }

        public bool Visibool { get; set; } = true;
        /// <summary>
        /// Called by excel to determine whether this ribbon is visible
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public bool OnVisible(IRibbonControl control)
        {
            //return false;
            return Visibool;
        }

        /// <summary>
        /// Generate xml for the ribbon
        /// </summary>
        /// <returns></returns>
        public string GenerateXml()
        {
            //Create a list of all methods with a ButtonAttribute
            var type = this.GetType();

            //Create a list of the property attributes
            var propertyAttributes = new List<RibbonElementAttribute>();
            foreach (var propInfo in type.GetProperties())
            {
                RibbonElementAttribute propAttribute = (RibbonElementAttribute)propInfo.GetCustomAttributes(typeof(RibbonElementAttribute), false).FirstOrDefault();
                if (propAttribute == null) continue;
                propAttribute.CallerName = propInfo.Name;
                propertyAttributes.Add(propAttribute);
            }

            //Create a list of ribbon element attributes
            //This is the list ordering which is used for the ribbon
            //So, the propertyAttributes are also inserted into this list using the get_PropertyName accessor method
            var ribbonElementAttributes = new List<RibbonElementAttribute>();
            foreach (var methInfo in type.GetMethods())
            {
                RibbonElementAttribute ribbonAttribute = (RibbonElementAttribute)methInfo.GetCustomAttributes(typeof(RibbonElementAttribute), false).FirstOrDefault();
                if (ribbonAttribute == null)
                {
                    if (methInfo.Name.StartsWith("get_"))
                    {
                        var propName = methInfo.Name.Substring(4);
                        var foundPropAttribute = propertyAttributes.Where(a => a.CallerName == propName).FirstOrDefault();
                        if(foundPropAttribute != null)
                        {
                            ribbonElementAttributes.Add(foundPropAttribute);

                        }
                    }
                    continue;
                }
                ribbonAttribute.CallerName = methInfo.Name;
                ribbonElementAttributes.Add(ribbonAttribute);
            }
            var x = type.GetMethods().ToList();

            //Group the ButtonAttribute objects by the GroupText
            var groupedByGroups = ribbonElementAttributes.GroupBy(b => b.GroupText);

            //Start generating the xml
            _GroupIndex = 0;
            _UIElementIndex = 0;

            //Create the ribbon xml by getting all the xml
            var sb = new StringBuilder();
            string ribbonXmlStart = $@"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad=""OnLoad"">
      <ribbon>
        <tabs>
          <tab id='tab1' label='{TabName}'>";
            sb.AppendLine(ribbonXmlStart);

            var groupText = new List<string>();
            foreach (var group in groupedByGroups)
            {
                sb.AppendLine($"            <group id='group{GetNextGroupIndex()}' label='{group.Key}'>");

                foreach (var but in group)
                {
                    sb.AppendLine($"                {but.InitAndGetXml(this)}");
                    //sb.AppendLine($"                <button id='button{GetNextButtonIndex()}' label='{but.ButtonText}' onAction='{but.MethodName}'/>");
                }

                sb.AppendLine($"            </group>");
            }

            string ribbonXmlEnd = @"
          </tab>
        </tabs>
      </ribbon>
    </customUI>";

            sb.AppendLine(ribbonXmlEnd);

            return sb.ToString();
        }

        #region Images

        /// <summary>
        /// A dictionary of the ID and the dynamic label object
        /// </summary>
        private Dictionary<string, Image> IconImages { get; set; } = new Dictionary<string, Image>();

        /// <summary>
        /// Called by the UI whenever it needs to get an image for a control, returns the control to be shown on the label
        /// </summary>
        public Image GetImage(IRibbonControl control)
        {
            if (IconImages.TryGetValue(control.Id, out Image img))
            {
                return img;
            }
            return null;
        }

        /// <summary>
        /// Registers the element so it's image can be retrieved by the UI
        /// Must be called by the element attribute when it is being initialised.
        /// </summary>
        public void RegisterIconImage(string id, Image img) => IconImages[id] = img;

        #endregion

        #region Dynamic labels

        /// <summary>
        /// A dictionary of the ID and the dynamic label object
        /// </summary>
        private Dictionary<string, DynamicText> DynamicLabels { get; set; } = new Dictionary<string, DynamicText>();

        /// <summary>
        /// Called by the UI whenever it needs to refresh a dynamic label button, returns the text to be shown on the label
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public string GetDynamicLabel(IRibbonControl control)
        {
            if(DynamicLabels.TryGetValue(control.Id, out DynamicText dl))
            {
                return dl.Text;
            }
            return "NO LABEL";
        }

        /// <summary>
        /// Registers the element so it's text can be retrieved by the UI
        /// Must be called by the element attribute when it is being initialised.
        /// </summary>
        public void RegisterDynamicLabel(string id, DynamicText value) => DynamicLabels[id] = value;

        #endregion

        #region Dynamic comboboxes

        private Dictionary<string, DynamicText> DynamicComboboxText { get; set; } = new Dictionary<string, DynamicText>();

        /// <summary>
        /// Called by the UI whenever it needs to refresh the combobox text, returns the selected text shown
        /// Callback: getText='...'   For control comboBox
        /// </summary>
        /// <param name="control"></param>
        public string GetDynamicComboboxText(IRibbonControl control)
        {
            if (DynamicComboboxText.TryGetValue(control.Id, out DynamicText dv))
            {
                return dv.Text;
            }
            return "NO TEXT";
        }

        
        /// <summary>
        /// Called by the UI whenever the user changes the combobox
        /// Callback: onChange='...'   For control comboBox
        /// </summary>
        /// <param name="control"></param>
        /// <param name="text"></param>
        public void ComboboxOnChange(IRibbonControl control, string text)
        {
            if (DynamicComboboxText.TryGetValue(control.Id, out DynamicText dv))
            {
                dv.Text = text;
            }
        }

        /// <summary>
        /// Registers the combobox so it's text can be retrieved by the UI
        /// Must be called by the combobox attribute when it is being initialised.
        /// </summary>
        public void RegisterDynamicDynamicComboboxText(string id, DynamicText value) => DynamicComboboxText[id] = value;

        #endregion

        #region Checkbox value

        /// <summary>
        /// A dictionary of the ID and the dynamic label object
        /// </summary>
        private Dictionary<string, DynamicBool> CheckboxValues { get; set; } = new Dictionary<string, DynamicBool>();

        /// <summary>
        /// Registers the element so it's text can be retrieved by the UI
        /// Must be called by the element attribute when it is being initialised.
        /// </summary>
        public void RegisterCheckboxValue(string id, DynamicBool value) => CheckboxValues[id] = value;

        /// <summary>
        /// Called by the UI whenever it needs to get the value of a checkbox
        /// Callback: getPressed='...'   For control checkBox
        /// </summary>
        public bool CheckboxGetPressed(IRibbonControl control)
        {
            if (CheckboxValues.TryGetValue(control.Id, out DynamicBool cv))
            {
                return cv.Value;
            }
            return false;
        }

        /// <summary>
        /// Called by the UI whenever a checkbox is pressed
        /// Callback: onAction='...'   For control checkBox
        /// </summary>
        public void CheckboxOnAction(IRibbonControl control, bool pressed)
        {
            if (CheckboxValues.TryGetValue(control.Id, out DynamicBool cv))
            {
                cv.Value = pressed;
            }
        }

        #endregion

        /// <summary>
        /// Helper method used to find properties by name in this class
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="name"></param>
        /// <returns></returns>
        public T FindProperty<T>(string name) where T : class
        {
            var thisType = this.GetType();
            var pInfo = thisType.GetProperty(name);
            return (T)pInfo.GetValue(this);
        }

    }

}
