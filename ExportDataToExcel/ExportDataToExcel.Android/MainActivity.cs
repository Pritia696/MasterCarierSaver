﻿using Android.App;
using Android.Content.PM;
using Android.OS;
using Plugin.Permissions;
using Xamarin.Forms;

namespace ExportDataToExcel.Droid
{
    [Activity(Label = "ReportHC", Icon = "@drawable/znac", Theme = "@style/MainTheme", MainLauncher = true, ConfigurationChanges = ConfigChanges.ScreenSize | ConfigChanges.Orientation)]
    public class MainActivity : global::Xamarin.Forms.Platform.Android.FormsAppCompatActivity
    {
        protected override void OnCreate(Bundle bundle)
        {
            TabLayoutResource = Resource.Layout.Tabbar;
            ToolbarResource = Resource.Layout.Toolbar;

            base.OnCreate(bundle);
            Window.SetStatusBarColor(Android.Graphics.Color.DarkGreen);
            Window.SetTitleColor(Android.Graphics.Color.DarkGreen);
            Plugin.CurrentActivity.CrossCurrentActivity.Current.Init(this, bundle);
            Forms.Init(this, bundle);
            LoadApplication(new App());
        }

        public override void OnRequestPermissionsResult(int requestCode, string[] permissions, Permission[] grantResults)
        {
            PermissionsImplementation.Current.OnRequestPermissionsResult(requestCode, permissions, grantResults);
            base.OnRequestPermissionsResult(requestCode, permissions, grantResults);
        }
    }
}