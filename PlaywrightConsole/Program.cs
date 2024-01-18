
/*---------------------------------------------------------------------
 * This application demonstrates how to log in to Microsoft Entra ID,
 * formerly known as Azure AD, with a user account with MFA enabled
 * using browser automation with Playwright.
 * 
 * Read more about this on my blog at
 * https://mikaberglund.com/how-to-do-mfa-login-with-playwright/
---------------------------------------------------------------------*/

//---------------------------------------------------------------------
// This sample application assumes that you specify 3 command line
// arguments when calling the application.
using Microsoft.Playwright;
using OtpNet;

var username = args[0];
var password = args[1];
var mfaSecret = args[2];
//---------------------------------------------------------------------

//---------------------------------------------------------------------
/*
 * Create the necessary objects we need to start running browser
 * automation. Please feel free to play around with the settings below
 * to find what you need. There's a lot to configure, but to keep
 * things simple and readable, I've just added a few things.
 */
var pw = await Playwright.CreateAsync();
var browser = await pw.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
{
    Headless = false
});
var context = await browser.NewContextAsync(new BrowserNewContextOptions
{
    Locale = "en-GB"
});
var page = await context.NewPageAsync();
await page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Screen });
//---------------------------------------------------------------------

var portalUrl = "https://www.microsoft365.com";

// Start off by navigating to the portal and force us to log in.
await page.GotoAsync($"{portalUrl}/login");

//---------------------------------------------------------------------
// Here we just define some commonly used CSS selectors.
var loginInputSelector = "input[type=email]";
var passwordInputSelector = "input[type=password]";
var submitSelector = "input[type=submit]";
var otpInputSelector = "input[name=otc]";
var kmsiCheckboxSelector = "#KmsiCheckboxField";
//---------------------------------------------------------------------

//---------------------------------------------------------------------
// Next we fill in the user name and passwords. Whenever we start a
// new browser context, we can rely on the fact that there are no
// cookies from a previous session, so we don't have to deal with
// the normal situation where your browser remembers your username from
// before. We can just write our code as if it was the very first time
// we are navigating to the site.
await page.FillAsync(loginInputSelector, username);
await page.ClickAsync(submitSelector); // The Next button

// Just to make our code more robust, wait for the username text box
// to be removed from the DOM before we continue with the password.
// This is because the login page is dynamically created with
// animations and stuff, and the CSS selectors have different meaning
// during different phases of the login process.
await page.Locator(loginInputSelector)
    .WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Detached });

// Now we do the password.
await page.FillAsync(passwordInputSelector, password);
await page.ClickAsync(submitSelector);

// And again, wait for the password text box to be removed from the DOM.
await page.Locator(passwordInputSelector)
    .WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Detached });

//---------------------------------------------------------------------

//---------------------------------------------------------------------
// Now we need to examine whether we need to do MFA or not. We do this
// by looking for two different elements, and which ever is found first
// determines whether MFA is required or not.
// If the element that is found first is the "Don't show this again"
// checkbox on the Keep me signed in page, then we don't need MFA, and
// the submit button is the final "Yes" button on the keep me signed
// in page.
// 
// But, if the first element is the one-time code input field, then
// we need to create a one-time password using the MFA secret.
var waiterIndex = Task.WaitAny(
    page.WaitForSelectorAsync(otpInputSelector),
    page.WaitForSelectorAsync(kmsiCheckboxSelector)
);

if(waiterIndex == 0)
{
    var totp = new Totp(Base32Encoding.ToBytes(mfaSecret));
    var otp = totp.ComputeTotp();
    await page.FillAsync(otpInputSelector, otp);
    await page.ClickAsync(submitSelector);

    // Wait until the OTP input box has been removed from the DOM.
    await page.Locator(otpInputSelector)
        .WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Detached });
}
//---------------------------------------------------------------------

//---------------------------------------------------------------------
// Now just finally click the "Yes" button on the Keep me signed page.
await page.ClickAsync(submitSelector);
//---------------------------------------------------------------------

//---------------------------------------------------------------------
// Now you have a page object that you have successfully logged in to
// the Microsoft 365 portal with using the account you specified as
// argument when launched the application.
// You can now continue to navigate to any Microsoft 365 service that
// the logged in user has access to and perform automation tasks.

//---------------------------------------------------------------------


Console.WriteLine("Press any key to quit.");
Console.ReadKey();

await context.DisposeAsync();
await browser.DisposeAsync();
pw.Dispose();
