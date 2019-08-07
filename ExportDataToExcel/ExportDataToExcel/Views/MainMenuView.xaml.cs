using ExportDataToExcel.Models;
using ExportDataToExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace ExportDataToExcel.Views
{
    public partial class MainMenuView : ContentPage
    {
        int t = 0;
        int t2 = 0;
        int t3 = 0;
        int t4 = 0;
        int t5 = 0;
        int t6 = 0;
        int t7 = 0;
        int t8 = 0;
        int t9 = 0;
        int t10 = 0;

        List<String> cars = new List<String>
            {
                "CAT 772G 07-88","CAT 772G 21-59","CAT 772G 21-60","CAT 773G 95-04","Volvo A40G 19-51","Volvo A40G 19-52",
                "Volvo A40F 74-01","Volvo A40F 74-02","Volvo A40E 66-29","Volvo A40E 66-30","Volvo A40E 66-31",
                "Volvo A40G 42-75","Volvo A40G 42-76","Volvo A40G 42-77","BelAZ 75-40 650","BelAZ 75-40 61-36","Volvo A40G 49-22","Volvo A40G 49-23"
            };
        List<String> drivers = new List<String>
            {
                "Aleshin A.","Antonov V.","Bahonov I.",
                "Davidov N.","Eliseev D.","Ermakov A.","Ermolov A.","Ezhikov S.",
                "Fedin A.","Fedorov S.","Filimonov A.","Hohlov V.","Isaev S.","Kabanov S.",
                "Karpuhin Y.","Kashirskiy I.","Krasavtsev E.","Krestyanskiy V.","Lobanov D.","Mastykash E.",
                "Nikitin I.E.","Nikitin R.V.","Nikulchev Y.","Novikov V.","Oskolkov V.","Panin V.",
                "Pleshikov A.","Pletnev Y.","Podolskiy I.N.","Podolskiy I.S.","Polosuev A.","Polzikov N.",
                "Rvantsev A.","Sergeev A.","Sidorkin A.","Sinitsin E.","Smolnyakov M.","Sorokin A.","Sverchkov A.",
                "Sychev D.","Tsikunov E.","Ulzutuev Y.","Zadonskiy V.","Zhelanov I."

            };
        List<String> technik = new List<String>
            {
                "Volvo EC480(42-18)"," CAT 374 (19-89)","CAT 374 (21-29)"," Liebherr 966 (03-91)",
                "Liebherr 976 (08-41)","Volvo L220F (46-52)","Volvo L220F (73-49)","Volvo L220F (86-22)"

            };
        public MainMenuView()
        {

            InitializeComponent();
            RegisterMesssages();

        }

        public void SavePhone(object sender, EventArgs e)
        {
            var model = GetReportModel();
            var t = new MainMenuViewModel(model);
            var res = t.ExportDataToExcelAsync(model);
        }

        public ReportModel GetReportModel()
        {
            var model = new ReportModel();
            model.Date = String.Format("{0}.{1}.{2}", date.SelectedItem, mounth.SelectedItem, year.SelectedItem);
            if (picker.SelectedItem != null)
            { model.WorkTime = picker.SelectedItem.ToString(); }
            model.MasterName = Family.SelectedItem.ToString();
            //считаем машины для первой техники
            
            var mashines1 = GetMashines(Grid2);
            var mashines2 = GetMashines(Grid3);
            var mashines3 = GetMashines(Grid4);
            var mashines4 = GetMashines(Grid5);
            var mashines5 = GetMashines(Grid6);
            var mashines6 = GetMashines(Grid7);
            var mashines7 = GetMashines(Grid8);
            var mashines8 = GetMashines(Grid9);
            var mashines9 = GetMashines(Grid10);
            var mashines10 = GetMashines(Grid11);
            model.Tecn = new List<Technique>();
            if (Teckn1.SelectedItem != null )
                
            //записываем первую технику 
            {
                if (Teckn1.SelectedItem.ToString() != "Удалить")
                {

               
                foreach (var mas in mashines1)
                {
                    mas.TechMins = new List<TechMin>();

                      var m = new TechMin
                    {
                        Name = Teckn1.SelectedItem.ToString(),
                        Index = model.Tecn.Count() + 1
                      };
                    mas.TechMins.Add(m);
                }

                var Tex1 = new Technique
                {
                    Id = 1,
                    DriverName = Fam1.SelectedItem.ToString(),
                    Name = Teckn1.SelectedItem.ToString(),
                    Poroda = Poroda1.SelectedItem.ToString(),
                    WorkPlace = Place1.SelectedItem.ToString(),
                    Mashines = mashines1,
                    Comment = comment1.Text
                };
                model.Tecn.Add(Tex1);
                }
                
            }

            if (Teckn2.SelectedItem != null)
            {
                if (Teckn2.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines2)
                    {
                        mas.TechMins = new List<TechMin>();

                        var m = new TechMin
                        {
                            Name = Teckn2.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex2 = new Technique
                    {
                        Id = 2,
                        DriverName = Fam2.SelectedItem.ToString(),
                        Name = Teckn2.SelectedItem.ToString(),
                        Poroda = Poroda2.SelectedItem.ToString(),
                        WorkPlace = Place2.SelectedItem.ToString(),
                        Mashines = mashines2,
                        Comment = comment2.Text

                    };
                    model.Tecn.Add(Tex2);
                }
            }

            if (Teckn3.SelectedItem != null)
            {
                if (Teckn3.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines3)
                    {
                        mas.TechMins = new List<TechMin>();

                        var m = new TechMin
                        {
                            Name = Teckn3.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex3 = new Technique
                    {
                        Id = 3,
                        DriverName = Fam3.SelectedItem.ToString(),
                        Name = Teckn3.SelectedItem.ToString(),
                        Poroda = Poroda3.SelectedItem.ToString(),
                        WorkPlace = Place3.SelectedItem.ToString(),
                        Mashines = mashines3,
                        Comment = comment3.Text

                    };
                    model.Tecn.Add(Tex3);
                }

            }
            if (Teckn4.SelectedItem != null)
            {
                if (Teckn4.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines4)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn4.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex4 = new Technique
                    {
                        Id = 4,
                        DriverName = Fam4.SelectedItem.ToString(),
                        Name = Teckn4.SelectedItem.ToString(),
                        Poroda = Poroda4.SelectedItem.ToString(),
                        WorkPlace = Place4.SelectedItem.ToString(),
                        Mashines = mashines4,
                        Comment = comment4.Text

                    };
                    model.Tecn.Add(Tex4);
                }
            }

            if (Teckn5.SelectedItem != null)
            {
                if (Teckn5.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines5)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn5.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex5 = new Technique
                    {
                        Id = 5,
                        DriverName = Fam5.SelectedItem.ToString(),
                        Name = Teckn5.SelectedItem.ToString(),
                        Poroda = Poroda5.SelectedItem.ToString(),
                        WorkPlace = Place5.SelectedItem.ToString(),
                        Mashines = mashines5,
                        Comment = comment5.Text

                    };
                    model.Tecn.Add(Tex5);
                }
            }
            if (Teckn6.SelectedItem != null)
            {
                if (Teckn6.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines6)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn6.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex6 = new Technique
                    {
                        Id = 6,
                        DriverName = Fam6.SelectedItem.ToString(),
                        Name = Teckn6.SelectedItem.ToString(),
                        Poroda = Poroda6.SelectedItem.ToString(),
                        WorkPlace = Place6.SelectedItem.ToString(),
                        Mashines = mashines6,
                        Comment = comment6.Text

                    };
                    model.Tecn.Add(Tex6);
                }

            }

            if (Teckn7.SelectedItem != null)
            {
                if (Teckn7.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines7)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn7.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex7 = new Technique
                    {
                        Id = 7,
                        DriverName = Fam7.SelectedItem.ToString(),
                        Name = Teckn7.SelectedItem.ToString(),
                        Poroda = Poroda7.SelectedItem.ToString(),
                        WorkPlace = Place7.SelectedItem.ToString(),
                        Mashines = mashines7,
                        Comment = comment7.Text

                    };
                    model.Tecn.Add(Tex7);
                }
            }

            if (Teckn8.SelectedItem != null)
            {
                if (Teckn8.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines8)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn8.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex8 = new Technique
                    {
                        Id = 8,
                        DriverName = Fam8.SelectedItem.ToString(),
                        Name = Teckn8.SelectedItem.ToString(),
                        Poroda = Poroda8.SelectedItem.ToString(),
                        WorkPlace = Place8.SelectedItem.ToString(),
                        Mashines = mashines8,
                        Comment = comment8.Text

                    };
                    model.Tecn.Add(Tex8);
                }

            }
            if (Teckn9.SelectedItem != null)
            {
                if (Teckn9.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines9)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn9.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex9 = new Technique
                    {
                        Id = 9,
                        DriverName = Fam9.SelectedItem.ToString(),
                        Name = Teckn9.SelectedItem.ToString(),
                        Poroda = Poroda9.SelectedItem.ToString(),
                        WorkPlace = Place9.SelectedItem.ToString(),
                        Mashines = mashines9,
                        Comment = comment9.Text

                    };
                    model.Tecn.Add(Tex9);
                }
            }
            if (Teckn10.SelectedItem != null)
            {
                if (Teckn10.SelectedItem.ToString() != "Удалить")
                {
                    foreach (var mas in mashines10)
                    {
                        mas.TechMins = new List<TechMin>();
                        var m = new TechMin
                        {
                            Name = Teckn10.SelectedItem.ToString(),
                            Index = model.Tecn.Count() + 1,
                        };
                        mas.TechMins.Add(m);
                    }
                    var Tex10 = new Technique
                    {
                        Id = 10,
                        DriverName = "",
                        Name = Teckn10.SelectedItem.ToString(),
                        Poroda = Poroda10.Text,
                        WorkPlace = "",
                        Mashines = mashines10,
                        Comment = comment10.Text

                    };
                    model.Tecn.Add(Tex10);
                }

            }
            model.TecnPlus = new List<TechniquePlus>();
            if (!String.IsNullOrEmpty(TexPlus1.Text))
            {
                var texPlus1 = new TechniquePlus
                {
                    Name = TexPlus1.Text,
                    DriverName= DriverPlus1.Text,
                    Work= WorkPlus1.Text,
                    Comment= comment11.Text
                };
                model.TecnPlus.Add(texPlus1);
            }
            if (!String.IsNullOrEmpty(TexPlus2.Text))
            {
                var texPlus2 = new TechniquePlus
                {
                    Name = TexPlus2.Text,
                    DriverName = DriverPlus2.Text,
                    Work = WorkPlus2.Text,
                    Comment = comment12.Text
                };
                model.TecnPlus.Add(texPlus2);
            }
            if (!String.IsNullOrEmpty(TexPlus3.Text))
            {
                var texPlus3 = new TechniquePlus
                {
                    Name = TexPlus3.Text,
                    DriverName = DriverPlus3.Text,
                    Work = WorkPlus3.Text,
                    Comment = comment13.Text
                };
                model.TecnPlus.Add(texPlus3);
            }
            if (!String.IsNullOrEmpty(TexPlus4.Text))
            {
                var texPlus4 = new TechniquePlus
                {
                    Name = TexPlus4.Text,
                    DriverName = DriverPlus4.Text,
                    Work = WorkPlus4.Text,
                    Comment = comment14.Text
                };
                model.TecnPlus.Add(texPlus4);
            }
            if (!String.IsNullOrEmpty(TexPlus5.Text))
            {
                var texPlus5 = new TechniquePlus
                {
                    Name = TexPlus5.Text,
                    DriverName = DriverPlus5.Text,
                    Work = WorkPlus5.Text,
                    Comment = comment15.Text
                };
                model.TecnPlus.Add(texPlus5);
            }

            return model;
        }

        public List<Mashine> GetMashines(Grid grid)
        {
            var counter = 0;
            var oldCounter = -1;
            var childreGrid2 = grid.Children.ToList();
            var mashines1 = new List<Mashine>();
            foreach (var ch in childreGrid2) //1table
            {
                if (counter != oldCounter)
                {
                    mashines1.Add(new Mashine());
                    oldCounter = counter;
                }

                var typeC = ch.GetType();

                if (typeof(Picker) == typeC)
                {
                    var r = (Xamarin.Forms.Picker)ch;
                    var value = r.SelectedItem;
                    mashines1[counter].Name = value.ToString();
                }
                if (typeof(Entry) == typeC)
                {
                    var r = (Entry)ch;
                    if (r.Placeholder == "Водитель")
                    {
                        mashines1[counter].DriverMName = r.Text;
                    }
                    if (r.Placeholder == "Рейсы")
                    {
                        mashines1[counter].Reis = r.Text;
                    }
                    if (r.Placeholder == "Плечо")
                    {
                        mashines1[counter].Plecho = r.Text;
                    }
                }
                if (typeof(Button) == typeC)
                {
                    counter++;
                }

            }
            return mashines1;
        }

        public void AddRow(object sender, EventArgs e)
        {
            Grid2.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker1 = new Picker
            {
            };

            picker1.ItemsSource = cars;
            var draiver1 = new Picker
            {
            };
            draiver1.ItemsSource = drivers;
            //var draiver1 = new Entry { Placeholder = "Водитель", FontSize = 14 };
            var reis1 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl1 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del1";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t);
            Grid2.Children.Add(picker1, 1, t);
            Grid2.Children.Add(draiver1, 2, t);
            Grid2.Children.Add(reis1, 3, t);
            Grid2.Children.Add(pl1, 4, t);
            Grid2.Children.Add(butDel, 5, t);
            t++;

        }

        public void AddRow2(object sender, EventArgs e)
        {
            Grid3.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del2";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t2);
            Grid3.Children.Add(picker2, 1, t2);
            Grid3.Children.Add(draiver2, 2, t2);
            Grid3.Children.Add(reis2, 3, t2);
            Grid3.Children.Add(pl2, 4, t2);
            Grid3.Children.Add(butDel, 5, t2);
            t2++;

        }
        public void AddRow3(object sender, EventArgs e)
        {
            Grid4.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker3 = new Picker
            {
            };

            picker3.ItemsSource = cars;
            var draiver3 = new Picker
            {
            };
            draiver3.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del3";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t3);
            Grid4.Children.Add(picker3, 1, t3);
            Grid4.Children.Add(draiver3, 2, t3);
            Grid4.Children.Add(reis2, 3, t3);
            Grid4.Children.Add(pl2, 4, t3);
            Grid4.Children.Add(butDel, 5, t3);
            t3++;
        }

        public void AddRow4(object sender, EventArgs e)
        {
            Grid5.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver3 = new Picker
            {
            };
            draiver3.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del4";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t4);
            Grid5.Children.Add(picker2, 1, t4);
            Grid5.Children.Add(draiver3, 2, t4);
            Grid5.Children.Add(reis2, 3, t4);
            Grid5.Children.Add(pl2, 4, t4);
            Grid5.Children.Add(butDel, 5, t4);
            t4++;
        }

        public void AddRow5(object sender, EventArgs e)
        {
            Grid6.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del5";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t5);
            Grid6.Children.Add(picker2, 1, t5);
            Grid6.Children.Add(draiver2, 2, t5);
            Grid6.Children.Add(reis2, 3, t5);
            Grid6.Children.Add(pl2, 4, t5);
            Grid6.Children.Add(butDel, 5, t5);
            t5++;
        }

        public void AddRow6(object sender, EventArgs e)
        {
            Grid7.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del6";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t6);
            Grid7.Children.Add(picker2, 1, t6);
            Grid7.Children.Add(draiver2, 2, t6);
            Grid7.Children.Add(reis2, 3, t6);
            Grid7.Children.Add(pl2, 4, t6);
            Grid7.Children.Add(butDel, 5, t6);
            t6++;
        }

        public void AddRow7(object sender, EventArgs e)
        {
            Grid8.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del7";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t7);
            Grid8.Children.Add(picker2, 1, t7);
            Grid8.Children.Add(draiver2, 2, t7);
            Grid8.Children.Add(reis2, 3, t7);
            Grid8.Children.Add(pl2, 4, t7);
            Grid8.Children.Add(butDel, 5, t7);
            t7++;
        }

        public void AddRow8(object sender, EventArgs e)
        {
            Grid9.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del8";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t8);
            Grid9.Children.Add(picker2, 1, t8);
            Grid9.Children.Add(draiver2, 2, t8);
            Grid9.Children.Add(reis2, 3, t8);
            Grid9.Children.Add(pl2, 4, t8);
            Grid9.Children.Add(butDel, 5, t8);
            t8++;
        }

        public void AddRow9(object sender, EventArgs e)
        {
            Grid10.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del9";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t9);
            Grid10.Children.Add(picker2, 1, t9);
            Grid10.Children.Add(draiver2, 2, t9);
            Grid10.Children.Add(reis2, 3, t9);
            Grid10.Children.Add(pl2, 4, t9);
            Grid10.Children.Add(butDel, 5, t9);
            t9++;
        }

        public void AddRow10(object sender, EventArgs e)
        {
            Grid11.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var picker2 = new Picker
            {
            };

            picker2.ItemsSource = cars;
            var draiver2 = new Picker
            {
            };
            draiver2.ItemsSource = drivers;
            var reis2 = new Entry { Placeholder = "Рейсы", FontSize = 14 };
            var pl2 = new Entry { Placeholder = "Плечо", FontSize = 14 };
            var butDel = new Button { Text = "-" };
            butDel.Clicked += DeleteRow;
            butDel.AutomationId = "del10";
            Grid.SetColumn(butDel, 5);
            Grid.SetRow(butDel, t10);
            Grid11.Children.Add(picker2, 1, t10);
            Grid11.Children.Add(draiver2, 2, t10);
            Grid11.Children.Add(reis2, 3, t10);
            Grid11.Children.Add(pl2, 4, t10);
            Grid11.Children.Add(butDel, 5, t10);
            t10++;
        }

        public void DeleteRow(object sender, EventArgs e)
        {
            try
            {
                var s = (Button)sender;
                var row = Grid.GetRow(s);
                var counter = t;
                var crid = Grid2;
                if (s.AutomationId == "del2")
                {
                    crid = Grid3;
                    counter = t2;
                }
                if (s.AutomationId == "del3")
                {
                    crid = Grid4;
                    counter = t3;
                }
                if (s.AutomationId == "del4")
                {
                    crid = Grid5;
                    counter = t4;
                }
                if (s.AutomationId == "del5")
                {
                    crid = Grid6;
                    counter = t5;
                }
                if (s.AutomationId == "del6")
                {
                    crid = Grid7;
                    counter = t6;
                }
                if (s.AutomationId == "del7")
                {
                    crid = Grid8;
                    counter = t7;
                }
                if (s.AutomationId == "del8")
                {
                    crid = Grid9;
                    counter = t8;
                }
                if (s.AutomationId == "del9")
                {
                    crid = Grid10;
                    counter = t9;
                }
                if (s.AutomationId == "del10")
                {
                    crid = Grid11;
                    counter = t10;
                }

                var children = crid.Children.ToList();
                foreach (var child in children.Where(c => Grid.GetRow(c) == row))
                {
                    crid.Children.Remove(child);
                }
                foreach (var child in children.Where(c => Grid.GetRow(c) > row))
                {
                    Grid.SetRow(child, Grid.GetRow(child) - 1);
                }
                var ee = Grid.GetRow(crid);
                crid.RowDefinitions.RemoveAt(counter - 1);
                if (s.AutomationId == "del2")
                {
                    t2--;
                }
                if (s.AutomationId == "del3")
                {
                    t3--;
                }
                if (s.AutomationId == "del4")
                {
                    t4--;
                }
                if (s.AutomationId == "del5")
                {
                    t5--;
                }
                if (s.AutomationId == "del6")
                {
                    t6--;
                }
                if (s.AutomationId == "del7")
                {
                    t7--;
                }
                if (s.AutomationId == "del8")
                {
                    t8--;
                }
                if (s.AutomationId == "del9")
                {
                    t9--;
                }
                if (s.AutomationId == "del10")
                {
                    t10--;
                }
                if (s.AutomationId == "del1")
                {
                    t--;
                }
            }
            catch (Exception exept)
            {
                var ex = exept;
            };

        }

    private void RegisterMesssages()
        {
            MessagingCenter.Subscribe<MainMenuViewModel>(this, "DataExportedSuccessfully", (m) =>
            {
                if (m != null)
                {
                    DisplayAlert("Info", "Data exported Successfully. The location is :" + m.FilePath, "OK");
                }
            });

            MessagingCenter.Subscribe<MainMenuViewModel>(this, "NoDataToExport", (m) =>
            {
                if (m != null)
                {
                    DisplayAlert("Warning !", "No data to export.", "OK");
                }
            });
        }

    }
}