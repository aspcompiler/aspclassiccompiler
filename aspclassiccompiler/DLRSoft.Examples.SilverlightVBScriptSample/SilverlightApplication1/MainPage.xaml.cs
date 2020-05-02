using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using Microsoft.Scripting.Hosting;
using Dlrsoft.VBScript.Hosting;

namespace SilverlightApplication1
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            
        }

        private void MyButton_Click(object sender, RoutedEventArgs e)
        {
            VbscriptHost vbscript = ((App)Application.Current).VBScript;
            ScriptScope scope = vbscript.CreateScope();
            scope.SetVariable("page", this);
            //this.MyButton.Content = "Clicked";
            string code = "page.MyButton.Content = \"Clicked\"";
            CompiledCode compiled = vbscript.Compile(code);
            compiled.Execute(scope);
        }
    }
}
