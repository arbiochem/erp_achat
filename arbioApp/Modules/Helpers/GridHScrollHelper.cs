using DevExpress.XtraGrid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Helpers
{
    public class GridHScrollHelper
    {
        GridView view;

        public GridHScrollHelper(GridView view)
        {
            this.view = view;
        }

        public void EnableScrollByColumns()
        {
            this.view.LeftCoordChanged += OnLeftCoordChanged;
        }

        void OnLeftCoordChanged(object sender, EventArgs e)
        {
            if (leftCoordChanged != 0) return;
            BeginScroll();
            try
            {
                GridColumnInfoArgs firstColumnInfo = GetFirstNonFixedColumnInfo();
                if (firstColumnInfo == null) return;
                int index = ViewInfo.ColumnsInfo.IndexOf(firstColumnInfo);
                if (IsRight)
                {
                    if (ScrollToNextColumn)
                    {
                        if (index < ViewInfo.ColumnsInfo.Count - 1)
                        {
                            GridColumnInfoArgs nextColumnInfo = ViewInfo.ColumnsInfo[++index];
                            SetLeftCoord(nextColumnInfo.Bounds.Right - Indent - 2, true);
                        }
                    }
                    else
                        SetLeftCoord(firstColumnInfo.Bounds.Right - Indent - 2, true);
                }
                else
                {
                    if (firstColumnInfo.Column.VisibleIndex != 0)
                        SetLeftCoord(firstColumnInfo.Bounds.Width - firstColumnInfo.Bounds.Right + Indent, false);
                    else
                        SetLeftCoord(view.LeftCoord, false);
                }
                oldLeftCoord = view.LeftCoord;
            }
            finally
            {
                EndScroll();
            }

        }
        int Indent
        {
            get { return ViewInfo.ViewRects.FixedLeft.Width + ViewInfo.ViewRects.IndicatorWidth; }
        }

        GridColumnInfoArgs GetFirstNonFixedColumnInfo()
        {
            GridColumnInfoArgs cArgs = null;
            foreach (GridColumnInfoArgs args in ViewInfo.ColumnsInfo)
                if (args.Column != null && args.Column.Fixed == DevExpress.XtraGrid.Columns.FixedStyle.None
                    && args.Bounds.Left <= Indent)
                    cArgs = args;
            return cArgs;
        }

        private void SetLeftCoord(int delta, bool isRight)
        {
            if (isRight)
                view.LeftCoord += delta;
            else
                view.LeftCoord -= delta;
        }

        int leftCoordChanged = 0, oldLeftCoord = 0;

        private void EndScroll()
        {
            leftCoordChanged--;
        }

        private void BeginScroll()
        {
            leftCoordChanged++;
        }

        GridViewInfo ViewInfo
        {
            get { return view.GetViewInfo() as GridViewInfo; }
        }

        bool IsRight
        {
            get { return view.LeftCoord > oldLeftCoord; }
        }

        bool ScrollToNextColumn
        {
            get
            {
                GridColumnInfoArgs args = GetFirstNonFixedColumnInfo();
                return args.Bounds.Right <= Indent;
            }
        }

    }
}
