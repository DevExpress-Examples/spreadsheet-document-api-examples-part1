using System;
using System.ComponentModel;
using DevExpress.XtraTreeList;
using DevExpress.Spreadsheet;

namespace SpreadsheetExamples {

    public class SpreadsheetNode {
        GroupsOfSpreadsheetExamples groups = new GroupsOfSpreadsheetExamples();
        GroupsOfSpreadsheetExamples owner;

        public SpreadsheetNode(string name) {
            Name = name;
        }
        [Browsable(false)]
        public GroupsOfSpreadsheetExamples Groups { get { return groups; } }
        public string Name { get; set; }

        [Browsable(false)]
        public GroupsOfSpreadsheetExamples Owner {
            get { return owner; }
            set { owner = value; }
        }
    }

    public class SpreadsheetExample : SpreadsheetNode {
        public SpreadsheetExample(string name, Action<Workbook> action) : base(name) {
            Action = action;
        }
        public Action<Workbook> Action { get; private set; }
    }

    public class GroupsOfSpreadsheetExamples : BindingList<SpreadsheetNode>, TreeList.IVirtualTreeListData {
        void TreeList.IVirtualTreeListData.VirtualTreeGetChildNodes(VirtualTreeGetChildNodesInfo info) {
            SpreadsheetNode obj = info.Node as SpreadsheetNode;
            info.Children = obj.Groups;
        }
        protected override void InsertItem(int index, SpreadsheetNode item) {
            item.Owner = this;
            base.InsertItem(index, item);
        }
        void TreeList.IVirtualTreeListData.VirtualTreeGetCellValue(VirtualTreeGetCellValueInfo info) {
            SpreadsheetNode obj = info.Node as SpreadsheetNode;
            switch (info.Column.Caption) {
                case "Name":
                    info.CellData = obj.Name;
                    break;
            }
        }
        void TreeList.IVirtualTreeListData.VirtualTreeSetCellValue(VirtualTreeSetCellValueInfo info) {
            SpreadsheetNode obj = info.Node as SpreadsheetNode;
            switch (info.Column.Caption) {
                case "Name":
                    obj.Name = (string)info.NewCellData;
                    break;
            }
        }
    }
}
