using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace InventorAddinLibrary
{

    #region General Documentation
    /*
    This is my first real attempt at creating Inventor API functionality. Frankly it is my first real
    software development project. Anyhow, this algorithim is a little complicated as least for me, 
    I have attempted to document the best I can.
    ----------------------------------------------------------
                    OVERALL FUNCTIONALITY
    ----------------------------------------------------------

    This functionality is pretty well documented within Autodesk Forums. However, the examples
    I found relied on on view names, or views to be placed in a certain order. This solution uses type checking 
    as well as determining postions to give the user more flexibility. Additionally. this functionality is coupled with the an addition function
    to arrange dimension placements.

    ----------------------------------------------------------
                           CREDITS
    ----------------------------------------------------------

    https://forums.autodesk.com/t5/inventor-forum/ilogic-and-scaling-view-size-automatically/td-p/7003140
    https://forums.autodesk.com/t5/inventor-ilogic-and-vb-net-forum/automatic-scale-and-view-position/m-p/5784680

    ----------------------------------------------------------
                         USER STORY
    ----------------------------------------------------------

    As a Inventor User, 
    I would like the ability to rescale and reposition drawing views with a click of a button, 
    so that I do not have to manually do so upon geometry change.

    ----------------------------------------------------------
                        PSEUDO CODE
    ----------------------------------------------------------

    1) We get a list of views from the active sheet.

    2) We iterate over the views to determine what type they are: Base, Projected, ISO, Flatpattern
        We add them to the drawingViewDictionary, with the first position of the dictionary value = the Name (Base, Projected, ISO, Flatpattern)
        We also assign the second position of the dictionary value to a horizontal or vertical array name.

    3) We then iterate over the viewDrawingDictionary, searching for view with the Name "Projected". 
        Then determine where the Projeced view locations are in respect to the Base view. 
        Reassigning the Projected name to a corresponding name e.g. Right, Left......
        We also assign the second position of the dictionary value to a horizontal or vertical array name.

    4) We itereate over the drawingViewDictionary determining the each of the views height and width if the scale was set to 1:1.
        We add the height and width values to the drawingViewDictionary values array.

    4) With the locations and 1:1 scale now known, we determine the required width and height to fit the drawings on the sheet.
        We calculate the scale required based on the the active sheet dimensions. Then iterate over views setting their scale.

    5) We itereate over the drawingViewDictionary determining the each of the views height and width if the scale set to the new calculated value.
        We add the height and width values to the drawingViewDictionary values array.

    6) We now use the new height and width values to determine where the view locations should be on the sheet. We iterate
        over the horizontal and vertical lists to determine the left most and bottom most view then set the corresponding views.
        We add these X,Y coordinates to the dictionary values array.

    7) Finally, we iterate over the drawingViewDictionary creating a 2D point from the X,Y coordinates we just added.
        Then assign the views center positon to that newly created point.

    ----------------------------------------------------------
                    ADDITIONAL NOTES:
    ----------------------------------------------------------

    The horizontalViewNames and vertivalViewNames arrays requirement thought process.
    In order to set the each views positions we have to determine what view is the left most or bottom most.
    Therefore the arrays are there so we interate of the drawingViewsDictionary in a certain order.
    Using the horizontal array for example we have {Left, Base, Right} when we search the drawingViewDictionary
    for Left we will figure out if it is first or not....

    For reference: The drawingViewDictionary format looks like
        DrawingView object as key
        List<String> {  ViewPosition, 
                        ViewPostionList, 
                        ViewHeight at 1:1, 
                        ViewWidth at 1:1, 
                        ViewHeight at FinalScale, 
                        ViewWidth at FinalScale, 
                        View X Position, 
                        View Y Position}
    */
    #endregion

    /// <summary>
    /// Class which represents rescaling and rearranging a drawing sheet functionality.
    /// </summary>
    public class ProcessDrawingSheet
    {

        #region Drawing Sheet Constants Definitions.
        /// <summary>
        /// Contstant that defines the desired white space in the X or horizontal direction UNITS = CM.
        /// </summary>
        private const double sheetAllowanceX = 6.0;

        /// <summary>
        /// Contstant that defines the desire white space in the Y or vertical direction UNITS = CM.
        /// </summary>
        private const double sheetAllowanceY = 5.0;

        /// <summary>
        /// Contstant that defines the gap in the horizonatal direction from left side of sheet UNITS = CM.
        /// </summary>
        private const double deviationX = 2.0;

        /// <summary>
        /// Contstant that defines the gap in the vertical direction from the bottom of the sheet UNITS = CM.
        /// </summary>
        private const double deviationY = 2.5;

        /// <summary>
        /// Contstant that defines the desired gap between each of the views UNITS = CM.
        /// </summary>
        private const double allowanceBetweenViews = 1.0;
        #endregion

        #region Global Variables

        /// <summary>
        /// The active drawing document to be processed.
        /// </summary>
        private Inventor.DrawingDocument oDrawDocument;

        /// <summary>
        /// The active Inventor appliction.
        /// </summary>
        private Inventor.Application oApplication;

        /// <summary>
        /// Array for iterating of views in the X (Horizontal) direction.
        /// </summary>
        private string[] horizontalViewNames = new string[] { "LeftView", "BaseView", "RightView" };

        /// <summary>
        /// Array for iterating of views in the X (Horizontal) direction.
        /// </summary>
        private string[] verticalViewNames = new string[] { "BottomView", "BaseView", "TopView", "FlatPatternView" };

        /// <summary>
        /// Dictionary that houses both a DrawingView Object as KEY
        /// As well as additional properties required for processing.
        /// List<String> {ViewPosition, ViewPostionList, ViewHeight at 1:1, ViewWidth at 1:1, ViewHeight at FinalScale, ViewWidth at FinalScale, ViewX Position, ViewYPosition}
        /// </summary>
        private Dictionary<DrawingView, List<string>> drawingViewDictionary = new Dictionary<DrawingView, List<string>>();

        #endregion

        /// <summary>
        /// Instantiates a new instance of the ProcessDrawingSheet class.
        /// </summary>
        /// <param name="document">Active Inventor DrawingDocument to be processed.</param>
        /// <param name="application">Active Inventor.Application instance.</param>
        public ProcessDrawingSheet(Inventor.Document document, Inventor.Application application)
        {
            this.oApplication = application;
            this.oDrawDocument = (DrawingDocument)document;

            this.InitialDrawingViews = oDrawDocument.ActiveSheet.DrawingViews;
            this.SetViewsTemporaryNames();
            this.SetProjectedViewsPositions();
            this.SetViewHeightWidthValues(this.BaseView.Scale);
            this.UpdateViewScale();
            this.SetViewHeightWidthValues(1);
            this.SetViewPostitionValues();
            this.UpdateViewPostions();
        }

        #region Properties Definitions

        /// <summary>
        /// Gets the active sheet height minus the view spacing allowances.
        /// </summary>
        public double AvailableSheetHeight
        {
            get
            {
                return this.oDrawDocument.ActiveSheet.Height - sheetAllowanceY - deviationY;
            }
        }

        /// <summary>
        /// Gets the active sheet width minus the view spacing allowances.
        /// </summary>
        public double AvailableSheetWidth
        {
            get
            {
                return this.oDrawDocument.ActiveSheet.Width - sheetAllowanceX - deviationX;
            }
        }

        /// <summary>
        /// Gets or Private Sets the base view of the drawing sheet.
        /// Is a view that has projected views from it, typically the front or top view
        /// of a part.
        /// </summary>
        public DrawingView BaseView { get; private set; }

        /// <summary>
        /// Gets the drawings Views upon intial call to method.
        /// </summary>
        public DrawingViews InitialDrawingViews { get; }

        /// <summary>
        /// Gets the Maximum scale size that can be uses.
        /// Return the smaller of the two betcause 1:4 and 1:2 ratios.
        /// </summary>
        public double MaxScale
        {
            get
            {
                double value = Math.Min(ScaleX, ScaleY);
                return value;
            }
        }

        /// <summary>
        /// Gets the sum height of View1 and View3 and View5 when set to 1:1 scale
        /// </summary>
        public double InititialMaxHeight
        {
            get
            {
                double value = 0;
                foreach (KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
                {
                    if (entry.Value[1].Contains("VerticalList"))
                    {
                        value += double.Parse(entry.Value[2]);
                    }

                }
                return value + ((this.ViewCountVertical - 1) * allowanceBetweenViews);
            }
        }

        /// <summary>
        /// Gets the sum width of View1 and View2 when set to 1:1 scale.
        /// </summary>
        public double InitialMaxWidth
        {
            get
            {
                double value = 0;
                foreach(KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
                {
                    if (entry.Value[1].Contains("HorizontalList"))
                    {
                        value += double.Parse(entry.Value[3]);
                    }
                        
                }
                return value + ((this.ViewCountHorizonal - 1) * allowanceBetweenViews);
            }
        }

        /// <summary>
        /// Gets the sum height of View1 and View3 and View5 at finaleScale
        /// </summary>
        public double MaxHeight
        {
            get
            {
                double value = 0;
                foreach (KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
                {
                    if (entry.Value[1].Contains("VerticalList"))
                    {
                        value += entry.Key.Height;
                    }

                }
                return value + ((this.ViewCountVertical - 1) * allowanceBetweenViews);
            }
        }

        /// <summary>
        /// Gets the sum width of View1 and View2 at final scale.
        /// </summary>
        public double MaxWidth
        {
            get
            {
                double value = 0;
                foreach (KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
                {
                    if (entry.Value[1].Contains("HorizontalList"))
                    {
                        value += entry.Key.Width;
                    }

                }
                return value + ((this.ViewCountHorizonal - 1) * allowanceBetweenViews);
            }
        }

        /// <summary>
        /// Gets the Scale in the X direction (Horizontal)
        /// </summary>
        public double ScaleX {
            get 
            {
                return this.AvailableSheetWidth / this.InitialMaxWidth;
            }
        }

        /// <summary>
        /// Gets the Scale in the Y direction (Vertical)
        /// </summary>
        public double ScaleY {
            get
            {
                return this.AvailableSheetHeight / this.InititialMaxHeight;
            }
        }

        /// <summary>
        /// Gets the count of views that are in the X direction (Horizontal)
        /// </summary>
        public int ViewCountHorizonal { get; private set; } = 0;

        /// <summary>
        /// Gets the count of views that are in the Y direction (Vertical)
        /// </summary>
        public int ViewCountVertical { get; private set; } = 0;


        #endregion

        /// <summary>
        /// Method which processes the views on the drawing sheet.
        /// Assigns Base View, IsoView, FlatpatternView, ProjectedView List Views based on their properties.
        /// </summary>
        private void SetViewsTemporaryNames()
        {
            foreach (DrawingView view in this.InitialDrawingViews)
            {
                if (view.IsFlatPatternView && view.ParentView == null && view.ViewType == DrawingViewTypeEnum.kStandardDrawingViewType)
                {
                    this.drawingViewDictionary.Add(view, new List<string> {"FlatPatternView", "VerticalList"});
                    this.ViewCountVertical++;
                }
                else if(view.IsFlatPatternView && view.ParentView != null)
                {
                    try
                    {
                        this.drawingViewDictionary.Remove(this.drawingViewDictionary.FirstOrDefault(x => x.Value[0] == "FlatPatternView").Key);
                    }
                    catch
                    {
                        // expecting errors after the first pass through the list.
                    }

                    if (this.BaseView == null)
                    {
                    this.drawingViewDictionary.Add(view.ParentView, new List<string> { "BaseView", "HorizontalList VerticalList" });
                    this.ViewCountHorizonal++;
                    }
                    this.drawingViewDictionary.Add(view, new List<string> { "Projected" });
                    this.BaseView = view.ParentView;
                }

                else if (view.ParentView != null)
                {
                    this.drawingViewDictionary.Add(view, new List<string> { "Projected" });

                    if (this.BaseView == null)
                    {
                        this.drawingViewDictionary.Add(view.ParentView, new List<string> { "BaseView", "HorizontalList VerticalList" });
                        this.ViewCountHorizonal++;
                        this.BaseView = view.ParentView;
                    }
                }
                else if (view.ViewType == DrawingViewTypeEnum.kProjectedDrawingViewType && view.ParentView == null)
                {
                    this.drawingViewDictionary.Add(view, new List<string> { "IsometricView", "NotApartOfList" });
                }
            }
        }

        /// <summary>
        /// Method which determines the current positioning of the projected views (the Base views orthographic projected views)
        /// Assigns corresponding DrawingView properies based on location from base view.
        /// Increments both the X and Y views counter.
        /// </summary>
        private void SetProjectedViewsPositions()
        {
            foreach(KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
            {
                if (entry.Value[0] == "Projected")
                {
                    // Had round the position values or would get miss readings.
                    double entryX = Math.Round(entry.Key.Position.X, 3);
                    double entryY = Math.Round(entry.Key.Position.Y, 3);

                    double baseX = Math.Round(this.BaseView.Position.X, 3);
                    double baseY = Math.Round(this.BaseView.Position.Y, 3);

                    if (entryX > baseX)
                    {
                        entry.Value[0] = "RightView";
                        entry.Value.Add("HorizontalList");
                        this.ViewCountHorizonal++;
                    }
                    else if (entryX < baseX)
                    {
                        entry.Value[0] = "LeftView";
                        entry.Value.Add("HorizontalList");
                        this.ViewCountHorizonal++;
                    }
                    else if (entryY > baseY)
                    {
                        entry.Value[0] = "TopView";
                        entry.Value.Add("VerticalList");
                        this.ViewCountVertical++;
                    }
                    else if (entryY < baseY)
                    {
                        entry.Value[0] = "BottomView";
                        entry.Value.Add("VerticalList");
                        this.ViewCountVertical++;
                    }
                }
            }
        }

        /// <summary>
        /// Method which iterates over the drawingViewDictionary adding
        /// height and width values to the array.
        /// </summary>
        /// <param name="scale">The scale factor used to calculate values.</param>
        private void SetViewHeightWidthValues(double scale)
        {
            foreach(KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
            {
                string heightValue = (entry.Key.Height * (1 / scale)).ToString();
                string widthValue = (entry.Key.Width *(1/scale)).ToString();
                entry.Value.Add(heightValue);
                entry.Value.Add(widthValue);
            }
        }

        /// <summary>
        /// Iterates over the views, sets there X,Y positions.
        /// </summary>
        private void SetViewPostitionValues()
        {
            double xInitialValue = (this.AvailableSheetWidth -MaxWidth) / (this.ViewCountHorizonal + 4);
            double yInitialValue = (this.AvailableSheetHeight - MaxHeight) / (this.ViewCountVertical + 1);
            double xPositionTracker = xInitialValue + deviationX;
            double yPositionTracker = yInitialValue + deviationY;

            // Iterate over the array, we add this step so we iterate in order to determine the left most view. 
            foreach (string horizontalName in this.horizontalViewNames)
            {
                if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(horizontalName)).Key != null)
                {
                    double viewWidth = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(horizontalName)).Key.Width;
                    xPositionTracker+= viewWidth / 2;
                    string value = xPositionTracker.ToString();
                    this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(horizontalName)).Value.Add(value);
                    xPositionTracker += xInitialValue + allowanceBetweenViews + viewWidth/2;
                }
            }

            // Iterate over the array, we add this step so we iterate in order to determine the bottom most view. 
            foreach (string verticalName in this.verticalViewNames)
            {
                if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(verticalName)).Key != null)
                {
                    double viewHeight = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(verticalName)).Key.Height;
                    yPositionTracker += viewHeight/2;
                    string value = yPositionTracker.ToString();
                    string baseX = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("BaseView")).Value[6];

                    if(verticalName != "BaseView")
                    {
                        this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(verticalName)).Value.Add(baseX);
                    }

                    this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(verticalName)).Value.Add(value);
                    yPositionTracker += yInitialValue + allowanceBetweenViews + viewHeight/2;
                }
            }

            // Set the horizontal views Y position based on the new value found from loop above.
            foreach (string horizontalName in this.horizontalViewNames)
            {
                if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(horizontalName)).Key != null)
                {
                    string baseY = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("BaseView")).Value[7];
                    this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains(horizontalName)).Value.Add(baseY);
                }
            }

            // Determine if there is an Isometric view, if so set its position.
            if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Key != null){
                double isoHeight = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Key.Height;
                double isoWidth = this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Key.Width;

                string isoX = (this.oDrawDocument.ActiveSheet.Width - isoWidth / 2 - 1).ToString();
                string isoY = (this.oDrawDocument.ActiveSheet.Height - isoHeight / 2 - 1).ToString();
                this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Value.Add(isoX);
                this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Value.Add(isoY);
            }
        }

        /// <summary>
        /// Method which iterates over the drawingViewDictionary
        /// creating a new Transient 2d Point Object then updating 
        /// the view to the corrected location.
        /// </summary>
        private void UpdateViewPostions()
        {
            foreach(KeyValuePair<DrawingView, List<string>> entry in this.drawingViewDictionary)
            {
                TransientGeometry otg = oApplication.TransientGeometry;

                Point2d point = otg.CreatePoint2d(double.Parse(entry.Value[6]), double.Parse(entry.Value[7]));

                entry.Key._Center = point;
            }
        }

        /// <summary>
        /// Method which iterates over the drawingViewDictionary setting
        /// each Key (DrawingView) scale to the desired scale
        /// </summary>
        private void UpdateViewScale()
        {
            double inc = 0;

            if(MaxScale < .01)
            {
                inc = .0001;
            }
            else if(MaxScale < .2)
            {
                inc = 0.01;
            }
            else if (MaxScale < 1)
            {
                inc = .1;
            }
            else
            {
                inc = 0.25;
            }

            double finalScale = Math.Floor(Math.Round(MaxScale, 4) / inc) * inc;

            if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("BaseView")).Key != null)
                this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("BaseView")).Key.Scale = finalScale;
            if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Key != null)
                this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("IsometricView")).Key.Scale = finalScale / 2;
            if (this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("FlatPatternView")).Key != null)
                this.drawingViewDictionary.FirstOrDefault(x => x.Value[0].Contains("FlatPatternView")).Key.Scale = finalScale;
        }
    }
}
