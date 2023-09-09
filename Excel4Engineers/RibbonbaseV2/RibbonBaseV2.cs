using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel4Engineers
{
    /// <summary>
    /// CONCEPT
    /// Draft - non-implemented code
    /// </summary>
    public class RibbonBaseV2
    {
        /*
         Use reflection to scan for GroupBase, then ControlBase etc...
        Intialise all the components and get thier xml
        Recursively calls child components.
         
         */
        public class Group1 : GroupBase
        {
            public class Button1 : ButtonBase
            {
                //Id is "Group1_Button1"
                public override void OnAction() => throw new NotImplementedException();
            }

            public class Button2 : ButtonBase
            {
                public override void OnAction() => throw new NotImplementedException();

            }

            public class Button3 : ButtonBase
            {
                public override void OnAction() => throw new NotImplementedException();
            }
        }


        #region Callback implementation example

        public Dictionary<string, IOnAction> OnActions;

        /// Callback: onAction='...'   For control button
        public void OnAction(IRibbonControl control)
        {
            if(OnActions.TryGetValue(control.Id, out IOnAction element))
            {
                element.OnAction();
            }
        }

        #endregion

        /*
         


        /// Callback: getDescription='...'   For control (several controls)
        public string GetDescription(IRibbonControl control)
        {
        }
        /// Callback: getEnabled='...'   For control (several controls)
        public bool GetEnabled(IRibbonControl control)
        {
        }
        /// Callback: getImage='...'   For control (several controls)
        public IPictureDisp GetImage(IRibbonControl control)
        {
        }
        /// Callback: getImageMso='...'   For control (several controls)
        public string GetImageMso(IRibbonControl control)
        {
        }
        /// Callback: getLabel='...'   For control (several controls)
        public string GetLabel(IRibbonControl control)
        {
        }
        /// Callback: getKeytip='...'   For control (several controls)
        public string GetKeytip(IRibbonControl control)
        {
        }
        /// Callback: getSize='...'   For control (several controls)
        public RibbonControlSize GetSize(IRibbonControl control)
        {
        }
        /// Callback: getScreentip='...'   For control (several controls)
        public string GetScreentip(IRibbonControl control)
        {
        }
        /// Callback: getSupertip='...'   For control (several controls)
        public string GetSupertip(IRibbonControl control)
        {
        }
        /// Callback: getVisible='...'   For control (several controls)
        public bool GetVisible(IRibbonControl control)
        {
        }
        /// Callback: getShowImage='...'   For control button
        public bool GetShowImage(IRibbonControl control)
        {
        }
        /// Callback: getShowLabel='...'   For control button
        public bool GetShowLabel(IRibbonControl control)
        {
        }
        /// Callback: onAction – repurposed='...'   For control button
        public void OnAction(IRibbonControl control, ref bool CancelDefault)
        {
        }

        /// Callback: getPressed='...'   For control checkBox
        public bool GetPressed(IRibbonControl control)
        {
        }
        /// Callback: onAction='...'   For control checkBox
        public void OnAction(IRibbonControl control, bool pressed)
        {
        }
        /// Callback: getItemCount='...'   For control comboBox
        public int GetItemCount(IRibbonControl control)
        {
        }
        /// Callback: getItemID='...'   For control comboBox
        public string GetItemID(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemImage='...'   For control comboBox
        public IPictureDisp GetItemImage(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemLabel='...'   For control comboBox
        public string GetItemLabel(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemScreenTip='...'   For control comboBox
        public string GetItemScreenTip(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemSuperTip='...'   For control comboBox
        public string GetItemSuperTip(IRibbonControl control, int index)
        {
        }
        /// Callback: getText='...'   For control comboBox
        public string GetText(IRibbonControl control)
        {
        }
        /// Callback: onChange='...'   For control comboBox
        public void OnChange(IRibbonControl control, string text)
        {
        }
        /// Callback: loadImage='...'   For control customUI
        public IPictureDisp LoadImage(string image_id)
        {
        }
        /// Callback: onLoad='...'   For control customUI
        public void OnLoad(IRibbonUI ribbon)
        {
        }
        /// Callback: getItemCount='...'   For control dropDown
        public int GetItemCount(IRibbonControl control)
        {
        }
        /// Callback: getItemID='...'   For control dropDown
        public string GetItemID(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemImage='...'   For control dropDown
        public IPictureDisp GetItemImage(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemLabel='...'   For control dropDown
        public string GetItemLabel(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemScreenTip='...'   For control dropDown
        public string GetItemScreenTip(IRibbonControl control, int index)
        {
        }
        /// Callback: getItemSuperTip='...'   For control dropDown
        public string GetItemSuperTip(IRibbonControl control, int index)
        {
        }
        /// Callback: getSelectedItemID='...'   For control dropDown
        public string GetSelectedItemID(IRibbonControl control)
        {
        }
        /// Callback: getSelectedItemIndex='...'   For control dropDown
        public int GetSelectedItemIndex(IRibbonControl control)
        {
        }
        /// Callback: onAction='...'   For control dropDown
        public void OnAction(IRibbonControl control, string selectedId, int selectedIndex)
        {
        }
        /// Callback: getContent='...'   For control dynamicMenu
        public string GetContent(IRibbonControl control)
        {
        }
        /// Callback: getText='...'   For control editBox
        public string GetText(IRibbonControl control)
        {
        }
        /// Callback: onChange='...'   For control editBox
        public void OnChange(IRibbonControl control, string text)
        {
        }
        /// Callback: getTitle='...'   For control menuSeparator
        public string GetTitle(IRibbonControl control)
        {
        }
        /// Callback: getPressed='...'   For control toggleButton
        public bool GetPressed(IRibbonControl control)
        {
        }
        /// Callback: onAction - repurposed='...'   For control toggleButton
        public void OnAction(IRibbonControl control, bool pressed, ref bool cancelDefault)
        {
        }
        /// Callback: onAction='...'   For control toggleButton
        public void OnAction(IRibbonControl control, bool pressed)
        {
        }
                 
         */
    }

    public abstract class RibbonElementBase
    {
        public List<RibbonElementBase> Children; //To be populated during initialisation

        public abstract void Initialise(RibbonBaseV2 ribbon, string id);

        public abstract string GetXml();
    }

    public abstract class GroupBase : RibbonElementBase
    {
        public override string GetXml()
        {
            throw new NotImplementedException();
        }

        public override void Initialise(RibbonBaseV2 ribbon, string id)
        {
            throw new NotImplementedException();
        }
    }

    public abstract class ControlBase : RibbonElementBase
    {
        
    }

    public abstract class ButtonBase : ControlBase, IOnAction
    {
        public override string GetXml()
        {
            throw new NotImplementedException();
        }

        public override void Initialise(RibbonBaseV2 ribbon, string id)
        {
            ribbon.OnActions.Add(id, this);
        }

        public abstract void OnAction();
    }

    public interface IOnAction
    {
        void OnAction();
    }
}
