using OutlookAddInMama;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace UnitTests
{
    
    
    /// <summary>
    ///Dies ist eine Testklasse für "ReplacementParserTest" und soll
    ///alle ReplacementParserTest Komponententests enthalten.
    ///</summary>
    [TestClass()]
    public class ReplacementParserTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Ruft den Testkontext auf, der Informationen
        ///über und Funktionalität für den aktuellen Testlauf bietet, oder legt diesen fest.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Zusätzliche Testattribute
        // 
        //Sie können beim Verfassen Ihrer Tests die folgenden zusätzlichen Attribute verwenden:
        //
        //Mit ClassInitialize führen Sie Code aus, bevor Sie den ersten Test in der Klasse ausführen.
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Mit ClassCleanup führen Sie Code aus, nachdem alle Tests in einer Klasse ausgeführt wurden.
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Mit TestInitialize können Sie vor jedem einzelnen Test Code ausführen.
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Mit TestCleanup können Sie nach jedem einzelnen Test Code ausführen.
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///Ein Test für "replaceAll"
        ///</summary>
        [TestMethod()]
        public void replaceAllTest()
        {
            Dictionary<string, string> replacementDictionary = new Dictionary<string, string>();
            replacementDictionary.Add("1", "you");
            replacementDictionary.Add("2", "$3");
            replacementDictionary.Add("3", "me");
            replacementDictionary.Add("10", "${1}");
            replacementDictionary.Add("1532", "he");
            replacementDictionary.Add("15321", "she");
            replacementDictionary.Add("namedgroup", "rose");

            ReplacementParser target = new ReplacementParser(replacementDictionary); // TODO: Passenden Wert initialisieren

            string input = "Hello $1!"; // TODO: Passenden Wert initialisieren
            string expected = "Hello you!"; // TODO: Passenden Wert initialisieren
            string actual;
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);
            
            input = "Hello ${1}!";
            expected = "Hello you!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello $${1}!";
            expected = "Hello ${1}!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello $$1!";
            expected = "Hello $1!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello ${1!";
            expected = "Hello ${1!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello ${1}1!";
            expected = "Hello you1!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello $11!";
            expected = "Hello $11!";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "Hello $1";
            expected = "Hello you";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1 are...";
            expected = "you are...";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "so $10";
            expected = "so ${1}";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1 are $1$1";
            expected = "you are youyou";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1 are $2";
            expected = "you are $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$a1 are $2";
            expected = "$a1 are $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1a are $2";
            expected = "youa are $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "${1a} are $2";
            expected = "${1a} are $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$15321a} is $2";
            expected = "shea} is $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$153212a} are $2";
            expected = "$153212a} are $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "${1532}1a} is $2";
            expected = "he1a} is $3";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1$$1$$$1${1}$${1}$$${1}$";
            expected = "you$1$youyou${1}$you$";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1+$$1+$$$1+${1}$+${1}+$$${1}$";
            expected = "you+$1+$you+you$+you+$you$";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = null;
            expected = null;
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$0";
            expected = "$0";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$1$2$3$4";
            expected = "you$3me$4";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "${namedgroup}$";
            expected = "rose$";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);

            input = "$namedgroup$";
            expected = "$namedgroup$";
            actual = target.replaceAll(input);
            Assert.AreEqual(expected, actual);
        }
    }
}
