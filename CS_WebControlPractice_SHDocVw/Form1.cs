using MSHTML;  // Microsoft.mshtml.dll
using SHDocVw;
using System;
using System.Windows.Forms;

namespace CS_WebControlPractice_SHDocVw
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            ShowIE();
        }

        private void ShowIE()
        {
            //var IE = new InternetExplorer();
            //var webBrowser = (IWebBrowserApp)IE;
            //webBrowser.Visible = true;
            //webBrowser.Navigate("http://www.google.com");

            SHDocVw.InternetExplorer IE = new SHDocVw.InternetExplorer();
            IWebBrowserApp webBrowser = (IWebBrowserApp)IE;

            dynamic url = "http://www.naver.com/";
            //url = "https://" + "dvkomiy" + ":" + "ta790909-1234" + "@sso.volkswagen.de/kpmweb/Index.action";
            url = "https://sso.volkswagen.de/kpmweb/Index.action";

            if (true)
            {
                url = "https://nid.naver.com/user2/V2Join.nhn?m=agree&lang=ko_KR";
                IE.Visible = true;
                IE.Navigate2(ref url); // 간단한 표현
                WaitPageLoading(IE);

                HTMLDocument doc = (HTMLDocument)IE.Document;
                IHTMLDocument2 doc2 = (IHTMLDocument2)IE.Document;
                PrintHTML(doc);

                IHTMLElement SelectedElement = doc.getElementById("chk_all");
                SelectedElement.click();
                WaitPageLoading(IE);
                
                SelectedElement = doc.getElementById("btnAgree");
                SelectedElement.click();
                WaitPageLoading(IE);

                SelectedElement = doc.getElementById("mm");
                var DropdownItems = (IHTMLElementCollection)SelectedElement.children;
                foreach (IHTMLElement elem in DropdownItems)
                {
                    System.Diagnostics.Debug.WriteLine(" elem.GetType() ===> " + elem.GetType().ToString());
                    System.Diagnostics.Debug.WriteLine(" tagName == " + elem.tagName);
                    System.Diagnostics.Debug.WriteLine(" innerText == " + elem.innerText);
                    System.Diagnostics.Debug.WriteLine(" OuterText == " + elem.outerText);
                    System.Diagnostics.Debug.WriteLine(" outerHTML == " + elem.outerHTML);
                    System.Diagnostics.Debug.WriteLine(" ==================================================");

                    string innertext = elem.innerText;
                    if (innertext == "1 ")
                    {
                        elem.setAttribute("selected", "selected");
                        //SelectedElement.setAttribute("aria-label", "elem.innerText");
                    }
                }


                SelectedElement = doc.getElementById("name");
                System.Diagnostics.Debug.WriteLine(" id == " + SelectedElement.id);
                System.Diagnostics.Debug.WriteLine(" innerText == " + SelectedElement.innerText);
                System.Diagnostics.Debug.WriteLine(" OuterText == " + SelectedElement.outerText);
                System.Diagnostics.Debug.WriteLine(" outerHTML == " + SelectedElement.outerHTML);
                System.Diagnostics.Debug.WriteLine(" tagName == " + SelectedElement.tagName);
                System.Diagnostics.Debug.WriteLine(" title == " + SelectedElement.title);

                if (SelectedElement.isTextEdit == true)
                {
                    SelectedElement.setAttribute("innerText", "Test");
                }

                

                HTMLInputElement HTMLInput;
                HTMLInput = doc2.activeElement as HTMLInputElement;
                var tag = HTMLInput.document as HTMLDocumentClass;
                //IHTMLElementCollection a = tag.getElementsByName("name");
                HTMLInput.value = "Test";
            }
            else
            {
                IE.Visible = true;
                IE.Navigate2(ref url); // 간단한 표현

                WaitPageLoading(IE);

                HTMLDocument doc = (HTMLDocument)IE.Document;
                PrintHTML(doc);

                IHTMLElement SelectedElement = doc.getElementById("problem_middle");
                SelectedElement.click();

                WaitPageLoading(IE);

                IHTMLInputElement textBox = (IHTMLInputElement)doc.getElementById("problemForm:panelBeschreibung:panelBodyBeschreibungen:beschreibungText:beschreibungText:beschreibungText:beschreibungText"); //lst-ib is the unique id of the textbox in which you want to set a value
                textBox.value = "Testing";

                IHTMLElementCollection elemcoll = doc.getElementById("erfassenForm:erfassenForm:workflowErfassungSelection:workflowErfassungSelection_input").children as IHTMLElementCollection;
                foreach (IHTMLElement elem in elemcoll)
                {
                    System.Diagnostics.Debug.WriteLine(" elem.GetType() ===> " + elem.GetType().ToString());
                    System.Diagnostics.Debug.WriteLine(" tagName == " + elem.tagName);
                    System.Diagnostics.Debug.WriteLine(" innerText == " + elem.innerText);
                    System.Diagnostics.Debug.WriteLine(" outerHTML == " + elem.outerHTML);
                    //System.Diagnostics.Debug.WriteLine(" children element 갯수 == " + elem.children.length); // text 제외한, tag 갯수만 포함. 
                    if (elem.getAttribute("id") != null)
                    {
                        System.Diagnostics.Debug.WriteLine(" id == " + elem.getAttribute("id"));
                    }

                    if (elem.innerText == "[21 MMI-H3 ] E/E")
                    {
                        elem.click();
                        elem.setAttribute("innerText", "[21 MMI-H3 ] E/E");

                    }
                    System.Diagnostics.Debug.WriteLine(" ==================================================");
                }

                SelectedElement = doc.getElementById("erfassenForm:erfassenForm:workflowErfassungSelection:workflowErfassungSelection_input");
                var DropdownItems = (IHTMLElementCollection)SelectedElement.children;

                var te = DropdownItems.length;
                System.Diagnostics.Debug.WriteLine(te);

                foreach (IHTMLElement option in DropdownItems)
                {
                    string temp = option.getAttribute("value").ToString();
                    System.Diagnostics.Debug.WriteLine(temp);

                    string temp2 = option.getAttribute("text").ToString();
                    System.Diagnostics.Debug.WriteLine(temp2);

                    if (temp == "3")
                    {
                        SelectedElement.setAttribute("value", "3");
                        option.setAttribute("selected", "selected");
                    }
                }
                SelectedElement.setAttribute("value", "4");


                IHTMLElementCollection SelectedElement_byName = doc.getElementsByName("erfassenForm:erfassenForm:workflowErfassungSelection:workflowErfassungSelection_input");

                //{
                //    var temp = item.getAttribute("value");
                //    System.Diagnostics.Debug.WriteLine(temp);
                //    temp = item.getAttribute("option");
                //    System.Diagnostics.Debug.WriteLine(temp);
                //    temp = item.getAttribute("text");
                //    System.Diagnostics.Debug.WriteLine(temp);

                //}
                foreach (IHTMLElement item in SelectedElement_byName)
                {
                    var temp = item.getAttribute("value");
                    System.Diagnostics.Debug.WriteLine(temp);
                    temp = item.getAttribute("option");
                    System.Diagnostics.Debug.WriteLine(temp);
                    temp = item.getAttribute("text");
                    System.Diagnostics.Debug.WriteLine(temp);
                }
                IHTMLElementCollection SelectedElement_nyTagname = doc.getElementsByTagName("erfassenForm:erfassenForm:workflowErfassungSelection:workflowErfassungSelection_input");
                IHTMLElementCollection SelectedElement_nyClassName = doc.getElementsByClassName("erfassenForm:erfassenForm:workflowErfassungSelection:workflowErfassungSelection_input");

            }




            IE.Quit();
        }

        private void WaitPageLoading(InternetExplorer IE)
        {
            while (IE.Busy == true || IE.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE)
            {
                System.Threading.Thread.Sleep(100);
            }
        }

        public void SetComboItem(string id, string value)
        {
            //HtmlElement ee = this.webBrowser1.Document.GetElementById(id);
            //foreach (HtmlElement item in ee.Children)
            //{
            //    if (item.OuterHtml.ToLower().IndexOf(value.ToLower()) >= 0)
            //    {
            //        item.SetAttribute("selected", "selected");
            //        item.InvokeMember("onChange");
            //    }
            //    else
            //    {
            //        item.SetAttribute("selected", "");
            //    }
            //}

            //ee = this.webBrowser1.Document.GetElementById(id + "-input");
            //ee.InnerText = value;
        }

        public void PrintHTML(HTMLDocument htmlDocument)
        {
            System.Diagnostics.Debug.WriteLine("doc.title == " + htmlDocument.title);
            System.Diagnostics.Debug.WriteLine("doc.URL == " + htmlDocument.url);
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
            System.Diagnostics.Debug.Write(htmlDocument.body.parentElement.outerHTML); // 모든 html 보여주기
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
        }
        public void PrintHTML(IHTMLDocument2 htmlDocument)
        {
            System.Diagnostics.Debug.WriteLine("doc.title == " + htmlDocument.title);
            System.Diagnostics.Debug.WriteLine("doc.URL == " + htmlDocument.url);
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
            System.Diagnostics.Debug.Write(htmlDocument.body.parentElement.outerHTML); // 모든 html 보여주기
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
        }

        public void PrintHTML(HtmlDocument htmlDocument)
        {
            System.Diagnostics.Debug.WriteLine("doc.title == " + htmlDocument.Title);
            System.Diagnostics.Debug.WriteLine("doc.URL == " + htmlDocument.Url);
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
            System.Diagnostics.Debug.Write(htmlDocument.Body.Parent.OuterHtml); // 모든 html 보여주기
            System.Diagnostics.Debug.WriteLine(" -------------------------------\n\n");
        }
    }
}