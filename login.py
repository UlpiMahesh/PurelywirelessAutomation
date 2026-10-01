from playwright.sync_api import Page


def login(page: Page, username: str, password: str) -> str:
    page.goto("https://www.t-mobiledealerordering.com/")

    page.fill("#userid", username)
    page.fill("#password", password)

    page.click("input[name='AgreeTerms']")
    page.click("a[name='login']")

    page.wait_for_load_state("load")
    page.wait_for_timeout(5000)

    print(f"[{username}] URL: {page.url}")

    page_text = page.locator("body").inner_text().lower()

    # --------------------------------------------------
    # 1. PASSWORD EXPIRED
    # --------------------------------------------------

    if (
        "change password" in page_text
        and "old password" in page_text
        and "new password" in page_text
    ):
        print(f"[{username}] ⚠️ PASSWORD EXPIRED")
        return "password_expired"

    # --------------------------------------------------
    # 2. INCORRECT USER ID / PASSWORD
    # --------------------------------------------------

    incorrect_message = (
        "your user id or password was incorrect. please try again."
    )

    if incorrect_message in page_text:
        print(f"[{username}] ❌ USER ID OR PASSWORD INCORRECT")
        return "password_incorrect"

    # --------------------------------------------------
    # 3. LOGIN FAILED
    # --------------------------------------------------

    if "login.do" in page.url:
        print(f"[{username}] ❌ LOGIN FAILED")
        return "login_failed"

    # --------------------------------------------------
    # 4. LOGIN SUCCESS
    # --------------------------------------------------

    print(f"[{username}] ✅ LOGIN SUCCESS")
    return "success"