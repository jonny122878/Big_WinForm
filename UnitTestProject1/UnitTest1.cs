//using System;
//using System.Linq;
//using System.Collections.Generic;
//using System.Windows.Forms;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using Rhino.Mocks;

//using WindowsFormsApp1.ViewModels;
//using WindowsFormsApp1.Algorithm;

//namespace UnitTestProject1
//{
//    [TestClass]
//    public class UnitTest1
//    {
//        [TestMethod]
//        public void TestMethod1()
//        {
//            #region Arrange
//            var validLoginViewModel = new ValidLoginViewModel
//            {
//                DisplayAuthentications = new List<Tuple<string, string, bool>>
//                {
//                    new Tuple<string, string, bool>("A","A1",true),
//                    new Tuple<string, string, bool>("A","A2",false),
//                    new Tuple<string, string, bool>("A","A3",false),
//                    new Tuple<string, string, bool>("A","A4",false),
//                    new Tuple<string, string, bool>("B","B1",false),
//                    new Tuple<string, string, bool>("B","B2",false),
//                    new Tuple<string, string, bool>("C","C1",true),
//                }
//            };
//            var excepted = 2;

//            var target = new AuthMenu();

//            #endregion

//            #region Act
//            //var menuStrip = MockRepository.GenerateStrictMock<MenuStrip>();
//            //menuStrip.Stub<MenuStrip>(r => r.Enabled = true);
//            var deny = MockRepository.GenerateStrictMock<ITest>();
//            //deny.Test = 5;
//            deny.Stub<ITest>(r => r.TestMethod(Arg<int>.Is.Anything)).Return(20);
//            var menuStrip = new MenuStrip();
//            target.LoadMenu(validLoginViewModel.DisplayAuthentications, menuStrip, deny);
//            var actual = target.Menus; 
//            #endregion

//            Assert.AreEqual(excepted, actual.Count);
//            //驗證出所產生的父和子類數量和bool是否正確

//            //Map出對應數量的ToolTipItem並且加入進去還有其對應事件

//            //實作對應View權限類別
//        }
//    }
//}
