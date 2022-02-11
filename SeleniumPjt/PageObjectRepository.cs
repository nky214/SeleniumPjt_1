using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumPjt
{
    internal class PageObjectRepository
    {
        private string URL = "https://rb-shoe-store.herokuapp.com/";
        private By emailInputBox = By.Id("remind_email_input");
        private By emailReturnMessageBox = By.XPath("//*[@id='flash']/div");
        private By emailSubmitButton = By.Id("remind_email_submit");
        private By promoCodeInputBox = By.Id("promo_code_input");
        private By promoSubmitButton = By.Id("promo_code_submit");
        private By welcomeMessage = By.XPath("/html/body/div[2]/div/h2");
        private By selectBrand = By.Id("brand");


        private By marchButton = By.XPath("//*[@id='header_nav']/nav/ul/li[3]/a");
        private By shoe_brand = By.ClassName("shoe_brand");
        private By shoe_name = By.ClassName("shoe_name");
        private By shoe_price = By.ClassName("shoe_price");
        private By shoe_description = By.ClassName("shoe_description");
        private By shoe_release_month = By.ClassName("shoe_release_month");

        private PageObjectRepository() { }
        private static PageObjectRepository instance = null;
        public static PageObjectRepository GetInstance()
        {

                if (instance == null)
                {
                    instance = new PageObjectRepository();
                }
                return instance;
            
        }

        public string GetURL()
        {
            return URL;
        }


        public By GetEmailInputBox()
        {
            return emailInputBox;
        }

        public By GetEmailReturnMessageBox()
        {
            return emailReturnMessageBox;
        }

        public By GetEmailSubmitButton()
        {
            return emailSubmitButton;
        }

        public By GetPromoCodeInputBox()
        {
            return promoCodeInputBox;
        }

        public By GetPromoSubmitButton()
        {
            return promoSubmitButton;
        }

        public By GetWelcomeMessage()
        {
            return welcomeMessage;
        }

        public By GetSelectBrand()
        {
            return selectBrand;
        }

        public By GetMarchButton()
        {
            return marchButton;
        }

        public By GetShoeBrand()
        {
            return shoe_brand;
        }

        public By GetShoeName()
        {
            return shoe_name;
        }

        public By GetShoePrice()
        {
            return shoe_price;
        }

        public By GetShoeDescription()
        {
            return shoe_description;
        }

        public By GetShoeReleaseMonth()
        {
            return shoe_release_month;
        }







    }
}
