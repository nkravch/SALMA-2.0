﻿#pragma checksum "..\..\..\View\MatrixReport.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "50D6A7A5AEE90EA2C158E9F0D26D364E"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.35312
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Controls.Ribbon;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Integration;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using WordToTFS.Properties;
using WordToTFS.View;


namespace WordToTFSWordAddIn.Views {
    
    
    /// <summary>
    /// MatrixReport
    /// </summary>
    public partial class MatrixReport : System.Windows.Window, System.Windows.Markup.IComponentConnector {
        
        
        #line 27 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListBox VerticalTypes;
        
        #line default
        #line hidden
        
        
        #line 36 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox StateVertical;
        
        #line default
        #line hidden
        
        
        #line 49 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button InsertButton;
        
        #line default
        #line hidden
        
        
        #line 58 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.Button CancelButton;
        
        #line default
        #line hidden
        
        
        #line 67 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ListBox HorizontalTypes;
        
        #line default
        #line hidden
        
        
        #line 75 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox StateHorisontal;
        
        #line default
        #line hidden
        
        
        #line 77 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.ComboBox Relateds;
        
        #line default
        #line hidden
        
        
        #line 86 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.CheckBox chkIncludeNotLinked;
        
        #line default
        #line hidden
        
        
        #line 94 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DatePicker dateFrom;
        
        #line default
        #line hidden
        
        
        #line 95 "..\..\..\View\MatrixReport.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.DatePicker dateTo;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/WordToTFS;component/view/matrixreport.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\View\MatrixReport.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            this.VerticalTypes = ((System.Windows.Controls.ListBox)(target));
            return;
            case 2:
            this.StateVertical = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 3:
            this.InsertButton = ((System.Windows.Controls.Button)(target));
            
            #line 56 "..\..\..\View\MatrixReport.xaml"
            this.InsertButton.Click += new System.Windows.RoutedEventHandler(this.InsertButtonClick);
            
            #line default
            #line hidden
            return;
            case 4:
            this.CancelButton = ((System.Windows.Controls.Button)(target));
            
            #line 65 "..\..\..\View\MatrixReport.xaml"
            this.CancelButton.Click += new System.Windows.RoutedEventHandler(this.CancelButtonClick);
            
            #line default
            #line hidden
            return;
            case 5:
            this.HorizontalTypes = ((System.Windows.Controls.ListBox)(target));
            return;
            case 6:
            this.StateHorisontal = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 7:
            this.Relateds = ((System.Windows.Controls.ComboBox)(target));
            return;
            case 8:
            this.chkIncludeNotLinked = ((System.Windows.Controls.CheckBox)(target));
            return;
            case 9:
            this.dateFrom = ((System.Windows.Controls.DatePicker)(target));
            
            #line 94 "..\..\..\View\MatrixReport.xaml"
            this.dateFrom.Loaded += new System.Windows.RoutedEventHandler(this.DatePicker_Loaded);
            
            #line default
            #line hidden
            return;
            case 10:
            this.dateTo = ((System.Windows.Controls.DatePicker)(target));
            
            #line 95 "..\..\..\View\MatrixReport.xaml"
            this.dateTo.Loaded += new System.Windows.RoutedEventHandler(this.DatePicker_Loaded);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

