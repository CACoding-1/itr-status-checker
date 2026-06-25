"""
itr_status_checker.pyw  —  Bulk Income-Tax Return Verification Status checker
=============================================================================
A single-file desktop tool (Tkinter GUI, no console window because of the
.pyw extension) that:

  1. Exports a blank Excel TEMPLATE  (Name | PAN | Password | AY | Status)
  2. Lets you UPLOAD the filled-in template
  3. RUNS the batch: for every row it
        logs in  ->  e-File  ->  Income Tax Returns  ->  View Filed Returns
        ->  Filter by the Assessment Year  ->  reads the LATEST status
        (e.g. "Processed with no demand/refund")  ->  records it in the
        Status column  ->  logs out  ->  next row, and repeats.
  4. Lets you DOWNLOAD an updated copy of the Excel with the Status column
     filled in (the original file is left untouched).

The Selenium login logic is the component you supplied (incometax_login.py),
embedded here so this stays a one-file deliverable.

-----------------------------------------------------------------------------
REQUIREMENTS  (install once, in a Command Prompt):
    pip install selenium openpyxl
    # Microsoft Edge must be installed (Selenium 4.6+ auto-manages the driver)

RUN:
    Double-click itr_status_checker.pyw   (or:  pythonw itr_status_checker.pyw)
-----------------------------------------------------------------------------
"""

import os
import time
import random
import threading
import queue
import re

# --- GUI -------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# --- Excel -----------------------------------------------------------------
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except Exception:                                            # pragma: no cover
    _OPENPYXL_OK = False

# --- Selenium --------------------------------------------------------------
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException


# =========================================================================== #
#  PART 1 — LOGIN COMPONENT  (from your incometax_login.py, lightly trimmed)   #
# =========================================================================== #
class LoginError(Exception):
    """Login could not be completed (e.g. Continue never proceeded)."""


class InvalidPasswordError(LoginError):
    """The portal reported an invalid/incorrect password. Do NOT retry."""


class IncomeTaxPortal:
    PORTAL_URL   = "https://www.incometax.gov.in/iec/foportal/"
    LOGIN_URL    = "https://eportal.incometax.gov.in/iec/foservices/#/login"
    PAGE_TIMEOUT = 20

    def __init__(self, driver=None, headless=False, log=None):
        self.log = log or (lambda m: print(m))
        self.driver = driver or self.build_driver(headless)

    # ---- driver construction --------------------------------------------- #
    @staticmethod
    def build_driver(headless=False):
        opts = EdgeOptions()
        opts.page_load_strategy = "none"   # never block on page loads; we poll
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        return webdriver.Edge(service=EdgeService(), options=opts)

    # ---- generic robust helpers ------------------------------------------ #
    def find(self, locators, timeout=None, need_clickable=False):
        timeout = self.PAGE_TIMEOUT if timeout is None else timeout
        end = time.time() + timeout
        while True:
            for by, val in locators:
                try:
                    els = self.driver.find_elements(by, val)
                except Exception:
                    els = []
                for el in els:
                    try:
                        if not el.is_displayed():
                            continue
                        if need_clickable and not el.is_enabled():
                            continue
                        return el
                    except Exception:
                        continue
            if time.time() >= end:
                raise TimeoutException(f"No locator matched in {timeout}s: {locators}")
            time.sleep(0.3)

    def click(self, locators, timeout=None):
        el = self.find(locators, timeout=timeout, need_clickable=True)
        try:
            el.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", el)
        return el

    def type(self, locators, text, timeout=None, slow=False):
        el = self.find(locators, timeout=timeout)
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        except Exception:
            pass
        try:
            el.click()
        except Exception:
            pass
        try:
            el.clear()
            if slow:
                for ch in str(text):
                    el.send_keys(ch)
                    time.sleep(random.uniform(0.06, 0.19))
            else:
                el.send_keys(text)
        except Exception:
            pass
        if (el.get_attribute("value") or "").strip() != str(text).strip():
            self.driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                el, str(text))
        return el

    def click_text(self, text, timeout=10, pick="shortest"):
        """Click an element by visible text. pick='shortest' (default),
        'top' (smallest y) or 'bottom' (largest y)."""
        js = r"""
        const want = arguments[0].trim().toLowerCase();
        const pick = arguments[1];
        const tags = ['button','a','input','span','label','li','div','td'];
        let best = null;
        for (const tag of tags){
          for (const el of document.getElementsByTagName(tag)){
            const r = el.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) continue;
            const txt = (el.innerText || el.value || '').trim().toLowerCase();
            if (!txt) continue;
            if (txt === want || txt.includes(want)){
              const cand = {el: el, len: txt.length, top: r.top};
              if (!best){ best = cand; }
              else if (pick === 'shortest' && cand.len < best.len){ best = cand; }
              else if (pick === 'top'      && cand.top < best.top){ best = cand; }
              else if (pick === 'bottom'   && cand.top > best.top){ best = cand; }
            }
          }
        }
        if (!best) return false;
        best.el.scrollIntoView({block:'center'});
        best.el.click();
        return true;
        """
        end = time.time() + timeout
        while time.time() < end:
            try:
                if self.driver.execute_script(js, text, pick):
                    return True
            except Exception:
                pass
            time.sleep(0.3)
        raise TimeoutException(f"No clickable element with text '{text}'")

    def check_box(self, timeout=8):
        def is_ticked():
            for b in self.driver.find_elements(By.XPATH, "//input[@type='checkbox']"):
                try:
                    if b.is_selected():
                        return True
                except Exception:
                    pass
            for b in self.driver.find_elements(
                    By.XPATH, "//*[@role='checkbox' or contains(@class,'p-checkbox-box')"
                              " or contains(@class,'mat-checkbox')]"):
                try:
                    cls = b.get_attribute("class") or ""
                    if (b.get_attribute("aria-checked") == "true"
                            or "p-highlight" in cls or "checked" in cls):
                        return True
                except Exception:
                    pass
            return False

        selectors = [
            "//label[contains(translate(.,'SECURE','secure'),'secure access')]",
            "//label[contains(translate(.,'AGREE','agree'),'agree')]",
            "//*[contains(@class,'p-checkbox-box')]",
            "//*[@role='checkbox']",
            "//*[contains(@class,'mat-checkbox-inner-container')]",
            "//label[.//input[@type='checkbox']]",
            "//input[@type='checkbox']",
        ]
        end = time.time() + timeout
        while time.time() < end:
            if is_ticked():
                return True
            for sel in selectors:
                for el in self.driver.find_elements(By.XPATH, sel):
                    try:
                        if not el.is_displayed():
                            continue
                        self.driver.execute_script(
                            "arguments[0].scrollIntoView({block:'center'});", el)
                        try:
                            el.click()
                        except Exception:
                            self.driver.execute_script("arguments[0].click();", el)
                        if is_ticked():
                            return True
                    except Exception:
                        continue
            for b in self.driver.find_elements(By.XPATH, "//input[@type='checkbox']"):
                try:
                    self.driver.execute_script(
                        "arguments[0].checked = true;"
                        "arguments[0].dispatchEvent(new Event('click',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", b)
                except Exception:
                    pass
            if is_ticked():
                return True
            time.sleep(0.3)
        return False

    def select_option(self, option_text, timeout=10):
        """Pick a value from whichever native <select> contains a matching
        option. Returns chosen text or None."""
        from selenium.webdriver.support.ui import Select
        want = str(option_text).strip().lower()
        end = time.time() + timeout
        while time.time() < end:
            for sel in self.driver.find_elements(By.TAG_NAME, "select"):
                try:
                    if not sel.is_displayed() or not sel.is_enabled():
                        continue
                    s = Select(sel)
                    idx = next((i for i, o in enumerate(s.options)
                                if want in (o.text or "").strip().lower()), None)
                    if idx is None:
                        continue
                    chosen = s.options[idx].text
                    try:
                        s.select_by_index(idx)
                    except Exception:
                        pass
                    self.driver.execute_script(
                        "arguments[0].selectedIndex = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", sel, idx)
                    try:
                        if want in (Select(sel).first_selected_option.text or "").strip().lower():
                            return chosen
                    except Exception:
                        return chosen
                except Exception:
                    continue
            time.sleep(0.3)
        return None

    def present(self, xpath):
        for el in self.driver.find_elements(By.XPATH, xpath):
            try:
                if el.is_displayed():
                    return True
            except Exception:
                pass
        return False

    def visible_text(self):
        try:
            return self.driver.execute_script("return document.body.innerText") or ""
        except Exception:
            return ""

    def wait_loaded(self, timeout=30, settle=0.6):
        """Block until the page has finished loading: document ready AND no
        visible loading spinner/overlay, held stable for `settle` seconds.
        Always returns within `timeout`, so a stuck spinner can never hang the
        run. Returns True if it reached an idle state, False on timeout."""
        busy_js = r"""
        if (document.readyState === 'loading') return true;
        const sels = ['.ngx-spinner-overlay','.mat-progress-spinner','.mat-spinner',
          '.p-progress-spinner','.spinner-border','[role=progressbar]',
          '[class*=spinner]','[class*=loader]','.loading-overlay'];
        for (const s of sels){
          let els;
          try { els = document.querySelectorAll(s); } catch(e){ continue; }
          for (const el of els){
            const r = el.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) continue;
            const st = window.getComputedStyle(el);
            if (st.visibility === 'hidden' || st.display === 'none' || st.opacity === '0') continue;
            return true;                 // a spinner/loader is visible
          }
        }
        return false;
        """
        end = time.time() + timeout
        idle_since = None
        while time.time() < end:
            try:
                busy = self.driver.execute_script(busy_js)
            except Exception:
                busy = False
            if busy:
                idle_since = None
            else:
                if idle_since is None:
                    idle_since = time.time()
                elif time.time() - idle_since >= settle:
                    return True
            time.sleep(0.25)
        return False

    # ---- THE LOGIN ------------------------------------------------------- #
    def login(self, pan, password, attempts=3, on_otp=None):
        d = self.driver
        pan = str(pan).strip().upper()
        self.log(f"Logging in — PAN {pan}")

        pan_locators = [
            (By.XPATH, "//input[contains(@placeholder,'PAN') or contains(@placeholder,'AADHAAR') or contains(@placeholder,'User ID')]"),
            (By.ID, "panAdharUserId"),
            (By.XPATH, "//input[@formcontrolname='userId']"),
            (By.XPATH, "//input[@type='text' and not(@disabled)]"),
        ]
        pw_locators = [
            (By.ID, "loginPasswordField"),
            (By.XPATH, "//input[@type='password']"),
            (By.XPATH, "//input[@formcontrolname='password']"),
        ]

        d.get(self.LOGIN_URL)
        try:
            self.type(pan_locators, pan, timeout=12)
        except Exception:
            self.log("  (direct login page didn't load — homepage fallback)")
            d.get(self.PORTAL_URL)
            try:
                self.click_text("Login", timeout=15)
            except Exception:
                pass
            self.type(pan_locators, pan, timeout=15)

        self.click([(By.XPATH, "//button[normalize-space()='Continue']"),
                    (By.ID, "continueBtn")])

        if self.check_box(timeout=8):
            self.log("  secure-access checkbox ticked")
        else:
            self.log("  WARNING: could not tick secure-access checkbox")

        self.type(pw_locators, password, slow=True)
        status = self._submit(password, pw_locators, attempts)
        if status == "invalid":
            raise InvalidPasswordError(f"Invalid password for PAN {pan}")
        if status == "failed":
            raise LoginError(f"Login did not proceed for PAN {pan}")

        time.sleep(1.5)
        page = d.page_source.lower()
        if "otp" in page or "captcha" in page:
            self.log("  OTP/CAPTCHA detected — solve it in the browser window")
            if on_otp:
                on_otp(d)
            else:
                # No console here (.pyw). Wait for the OTP screen to clear.
                self._wait_otp_cleared(timeout=180)

        self.wait_loaded(timeout=30)
        self.log("  login complete")
        return True

    def _wait_otp_cleared(self, timeout=180):
        """Poll until the OTP/CAPTCHA screen disappears (user solved it) or the
        dashboard loads. Used instead of input() because .pyw has no console."""
        end = time.time() + timeout
        while time.time() < end:
            low = self.visible_text().lower()
            if ("otp" not in low and "captcha" not in low) or "e-file" in low or "dashboard" in low:
                return True
            time.sleep(1.0)
        return False

    def _submit(self, password, pw_locators, attempts=3):
        continue_locs = [
            (By.XPATH, "//button[normalize-space()='Continue']"),
            (By.ID, "submitBtn"),
            (By.XPATH, "//button[contains(.,'Continue')]"),
        ]
        INVALID = ["invalid password", "incorrect password", "password is incorrect",
                   "invalid user id or password", "user id or password",
                   "password you entered", "wrong password"]
        REAUTH = ["request not authenticated"]

        for attempt in range(1, attempts + 1):
            if attempt > 1:
                self.log(f"  re-entering password (attempt {attempt}/{attempts})")
                self._erase_password(pw_locators)
                self.type(pw_locators, password, slow=True, timeout=8)
            try:
                self.click(continue_locs, timeout=8)
            except Exception:
                self.log(f"  WARNING: Continue not clickable (attempt {attempt})")

            end = time.time() + 8
            while time.time() < end:
                self._dismiss_login_here()
                low = self.visible_text().lower()
                if any(k in low for k in INVALID):
                    return "invalid"
                if not self.present("//input[@type='password']"):
                    return "ok"
                if any(k in low for k in REAUTH):
                    self.log("  'Request not authenticated' — will erase & retry")
                    break
                time.sleep(0.4)
        return "ok" if not self.present("//input[@type='password']") else "failed"

    def _erase_password(self, pw_locators):
        try:
            el = self.find(pw_locators, timeout=8)
        except Exception:
            return
        for action in (
            lambda: (el.click(), el.send_keys(Keys.CONTROL, "a"), el.send_keys(Keys.DELETE)),
            lambda: el.clear(),
            lambda: self.driver.execute_script(
                "arguments[0].value='';"
                "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el),
        ):
            try:
                action()
            except Exception:
                pass

    def _dismiss_login_here(self):
        try:
            return self.click_text("Login here", timeout=1)
        except Exception:
            return False

    def quit(self):
        try:
            self.driver.quit()
        except Exception:
            pass


# =========================================================================== #
#  PART 2 — WORKFLOW  (navigate, filter by AY, read status, log out)           #
# =========================================================================== #

# Known portal statuses, most-specific first. The LATEST status is whichever
# of these sits at the TOP of the filing's status timeline.
KNOWN_STATUSES = [
    "Processed with no demand/refund",
    "Processed with no demand / refund",
    "Processed with no demand or refund",
    "Processed with refund due",
    "Processed with demand due",
    "Processed with refund",
    "Processed with demand",
    "Refund Issued",
    "Refund Failed",
    "Refund Re-issued",
    "Successfully e-Verified",
    "Submitted and pending for e-Verification",
    "Pending for e-Verification",
    "Pending for verification",
    "Under Processing",
    "ITR Processed",
    "Return uploaded",
    "Defective",
    "Invalidated",
    "Transferred to Jurisdictional Assessing Officer",
    "Rectification processed",
    "Case transferred to AO",
]


def go_to_filed_returns(portal):
    """e-File -> Income Tax Returns -> View Filed Returns."""
    portal.log("  navigating: e-File > Income Tax Returns > View Filed Returns")
    portal.wait_loaded(timeout=30)
    portal.click_text("e-File", timeout=20)
    time.sleep(0.8)                       # let the menu expand
    portal.click_text("Income Tax Returns", timeout=20)
    time.sleep(0.8)
    portal.click_text("View Filed Returns", timeout=20)
    portal.wait_loaded(timeout=30)
    # Confirm the results page actually rendered.
    end = time.time() + 25
    while time.time() < end:
        low = portal.visible_text().lower()
        if "filed returns" in low or "filing type" in low or "acknowledgement" in low:
            return True
        time.sleep(0.5)
    portal.log("  WARNING: 'View Filed Returns' page did not clearly load")
    return False


def filter_by_ay(portal, ay):
    """Filter the list: click Filter -> open the Assessment Year dropdown ->
    pick the year -> click the panel's Filter button to apply."""
    ay = normalize_ay(ay)
    portal.log(f"  filtering by Assessment Year {ay}")
    portal.wait_loaded(timeout=20)

    # 1) Open the Filter panel.
    try:
        portal.click_text("Filter", timeout=10)
    except Exception:
        portal.log("  WARNING: could not find the Filter button")
    time.sleep(0.8)
    portal.wait_loaded(timeout=15)

    # 2) Open the Assessment Year dropdown and pick the year.
    if _choose_ay(portal, ay):
        portal.log(f"  Assessment Year set to {ay}")
    else:
        portal.log(f"  WARNING: could not set Assessment Year to {ay}")
    time.sleep(0.5)

    # 3) Apply — click the panel's blue 'Filter' button. Match it among real
    #    buttons by EXACT text 'Filter' (so 'Filter By' is ignored) and take
    #    the bottom-most one (the apply button sits below the toggle).
    apply_js = r"""
    let best = null;
    const els = document.querySelectorAll(
        'button,[role=button],input[type=button],input[type=submit],a');
    for (const el of els){
      const r = el.getBoundingClientRect();
      if (r.width === 0 || r.height === 0) continue;
      const txt = (el.innerText || el.value || '').trim().toLowerCase();
      if (txt !== 'filter') continue;                 // exact, not 'filter by'
      if (!best || r.top > best.top){ best = {el: el, top: r.top}; }
    }
    if (!best) return false;
    best.el.scrollIntoView({block:'center'});
    best.el.click();
    return true;
    """
    applied = False
    end = time.time() + 8
    while time.time() < end:
        try:
            if portal.driver.execute_script(apply_js):
                applied = True
                break
        except Exception:
            pass
        time.sleep(0.4)
    if not applied:
        portal.log("  WARNING: could not click the Filter (apply) button")
    portal.wait_loaded(timeout=30)


def _choose_ay(portal, ay):
    """Open the Assessment Year dropdown box and click the matching year.
    Tries a native <select> first, then the click-to-open custom dropdown."""
    d = portal.driver

    # Simplest case: a real <select>.
    if portal.select_option(ay, timeout=3):
        return True

    # Open the dropdown box that sits just below the 'Assessment Year' label.
    open_js = r"""
    const lbl = arguments[0].trim().toLowerCase();
    let label = null;
    for (const el of document.querySelectorAll('label,span,div,p')){
      if ((el.textContent||'').trim().toLowerCase() === lbl){ label = el; break; }
    }
    if (!label){
      for (const el of document.querySelectorAll('label,span,div,p')){
        if ((el.textContent||'').trim().toLowerCase().startsWith(lbl)){ label = el; break; }
      }
    }
    if (!label) return false;
    const lr = label.getBoundingClientRect();
    const sel = 'select,.p-dropdown,[role=combobox],.mat-select,.ui-dropdown,'
              + '.dropdown-toggle,.p-dropdown-trigger,div[class*=dropdown],'
              + 'div[class*=select],input[readonly]';
    let best = null;
    for (const el of document.querySelectorAll(sel)){
      const r = el.getBoundingClientRect();
      if (r.width===0 || r.height===0) continue;
      const dy = r.top - lr.top;
      if (dy < -5 || dy > 120) continue;            // at / just below the label
      const score = dy + Math.abs(r.left - lr.left) * 0.2;
      if (!best || score < best.score){ best = {el: el, score: score}; }
    }
    if (!best) return false;
    best.el.scrollIntoView({block:'center'});
    best.el.click();
    return true;
    """

    # Click the year option. The open option sits BELOW the closed trigger,
    # so among exact-text matches we pick the one with the largest 'top'.
    pick_js = r"""
    const want = arguments[0].trim().toLowerCase().replace(/\s+/g,'');
    let cands = [];
    for (const el of document.querySelectorAll(
        'li,[role=option],.p-dropdown-item,.mat-option,option,span,div,a')){
      if (el.children && el.children.length > 2) continue;     // leaf-ish only
      const r = el.getBoundingClientRect();
      if (r.width===0 || r.height===0) continue;
      const t = (el.textContent || el.value || '').trim().toLowerCase().replace(/\s+/g,'');
      if (t === want){ cands.push({el: el, top: r.top}); }
    }
    if (!cands.length) return false;
    cands.sort((a,b) => b.top - a.top);
    cands[0].el.scrollIntoView({block:'center'});
    cands[0].el.click();
    return true;
    """

    for _ in range(3):
        try:
            d.execute_script(open_js, "Assessment Year")
        except Exception:
            pass
        time.sleep(0.9)
        try:
            if d.execute_script(pick_js, ay):
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


# Phrases that mean the portal returned no filing for the chosen AY.
EMPTY_HINTS = [
    "no record", "no data", "no return", "not filed", "no filing",
    "data not available", "no result", "0 filing",
]


def read_latest_status(portal, settle=8):
    """Return the LATEST status for the filtered AY, or 'ITR not filed' when no
    return is shown. The latest status is the top-most label in the timeline."""
    portal.wait_loaded(timeout=20)
    # Report ONLY the status phrase (no dates / acknowledgement numbers): we
    # return the matched known status, not the element's raw text.
    scan_js = r"""
    const known = arguments[0];                 // original-case phrases
    let best = null;            // smallest 'top' = highest on the page = latest
    for (const el of document.querySelectorAll('span,div,p,td,li,strong,b,h3,h4')){
      const r = el.getBoundingClientRect();
      if (r.width===0 || r.height===0) continue;
      const raw = (el.innerText || el.textContent || '').trim();
      if (!raw || raw.length > 120) continue;
      const low = raw.toLowerCase();
      for (const k of known){
        if (low.includes(k.toLowerCase())){
          if (!best || r.top < best.top){ best = {text: k, top: r.top}; }
          break;
        }
      }
    }
    return best ? best.text : "";
    """
    end = time.time() + settle
    low = ""
    while time.time() < end:
        try:
            txt = portal.driver.execute_script(scan_js, KNOWN_STATUSES)
        except Exception:
            txt = ""
        if txt and txt.strip():
            return txt.strip()
        low = portal.visible_text().lower()
        if any(h in low for h in EMPTY_HINTS):
            return "ITR not filed"
        time.sleep(0.5)

    # Settled with no status text. A filing card always carries an
    # acknowledgement number, so its absence means nothing was filed.
    if "acknowledgement" in low:
        return "Status not found"
    return "ITR not filed"


def logout(portal):
    """Open the top-right profile menu and click 'Log Out' (matches the
    'My Profile / Change Password / Log Out' dropdown)."""
    portal.wait_loaded(timeout=15)
    d = portal.driver

    # Click the menu item whose EXACT text is Log Out / Logout / Sign Out.
    click_logout_js = r"""
    const wants = ['log out', 'logout', 'sign out', 'log off'];
    let best = null;
    for (const el of document.querySelectorAll('a,button,li,span,div,p')){
      if (el.children && el.children.length > 1) continue;        // leaf node only
      const r = el.getBoundingClientRect();
      if (r.width===0 || r.height===0) continue;
      const t = (el.innerText || el.textContent || '').trim().toLowerCase();
      if (wants.includes(t)){
        if (!best || t.length < best.len){ best = {el: el, len: t.length}; }
      }
    }
    if (!best) return false;
    best.el.scrollIntoView({block:'center'});
    best.el.click();
    return true;
    """

    for _ in range(3):
        _open_profile_menu(portal)
        time.sleep(1.0)
        try:
            if d.execute_script(click_logout_js):
                portal.wait_loaded(timeout=15)
                portal.log("  logged out")
                return True
        except Exception:
            pass
        time.sleep(0.6)
    portal.log("  (could not click Log Out — closing the browser ends the session)")
    return False


def _open_profile_menu(portal):
    """Click the profile toggle (avatar / name / caret) in the top-right header
    so the My Profile / Change Password / Log Out menu drops down."""
    open_js = r"""
    const W = window.innerWidth;
    function clsOf(el){
      const c = el.className;
      if (c && typeof c === 'object' && 'baseVal' in c) return c.baseVal;   // SVG
      return (c || '') + '';
    }
    let best = null;
    for (const el of document.querySelectorAll('img,i,svg,span,div,a,button,li')){
      const r = el.getBoundingClientRect();
      if (r.width===0 || r.height===0) continue;
      if (r.top > 110 || r.left < W*0.5) continue;          // top-right header band
      const c = clsOf(el).toLowerCase();
      const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
      let score = 0;
      if (el.tagName === 'IMG') score += 3;                                       // avatar
      if (/chevron|caret|angle|arrow|expand|dropdown|down/.test(c)) score += 4;   // caret icon
      if (/individual|\bhuf\b|profile|account/.test(txt) && txt.length < 40) score += 2;
      if (score === 0) continue;
      const s = score - (r.width * r.height) / 200000;       // prefer small toggles/icons
      if (!best || s > best.s){ best = {el: el, s: s}; }
    }
    if (!best){                                              // fallback: right-most clickable
      for (const el of document.querySelectorAll('a,button,img,i,span,div')){
        const r = el.getBoundingClientRect();
        if (r.width===0 || r.height===0 || r.top > 110) continue;
        if (!best || r.left > best.left){ best = {el: el, left: r.left}; }
      }
    }
    if (!best) return false;
    best.el.scrollIntoView({block:'center'});
    best.el.click();
    return true;
    """
    try:
        portal.driver.execute_script(open_js)
    except Exception:
        pass


# ---- small file/text utilities ------------------------------------------- #
def normalize_ay(ay):
    """'AY 2026-27' / '2026-27' / '2026-2027' -> '2026-27'."""
    s = str(ay).strip()
    s = re.sub(r"(?i)^a\.?y\.?\s*", "", s).strip()
    m = re.search(r"(20\d{2})\s*[-/]\s*(\d{2,4})", s)
    if m:
        start, end = m.group(1), m.group(2)
        if len(end) == 4:
            end = end[-2:]
        return f"{start}-{end}"
    return s


# =========================================================================== #
#  PART 3 — EXCEL  (template export + read rows + export updated copy)         #
# =========================================================================== #
HEADERS = ["Name", "PAN", "Password", "AY", "Status"]


def export_template(path):
    """Write a blank, formatted template workbook to `path`."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Returns"

    head_fill = PatternFill("solid", fgColor="1F4E78")
    head_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # A sample row (greyed) to show the expected format.
    sample = ["John Doe", "ABCDE1234F", "YourPassword", "2026-27", ""]
    for c, v in enumerate(sample, start=1):
        cell = ws.cell(row=2, column=c, value=v)
        cell.font = Font(italic=True, color="9C9C9C")
        cell.border = border

    widths = [26, 16, 18, 12, 34]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"

    # Instructions sheet.
    info = wb.create_sheet("Instructions")
    notes = [
        "HOW TO USE THIS TEMPLATE",
        "",
        "1. Fill ONE row per return you want to check.",
        "2. Columns:",
        "      Name      - any label to identify the taxpayer (for your reference only)",
        "      PAN       - 10-character PAN, e.g. ABCDE1234F",
        "      Password  - the e-Filing portal login password for that PAN",
        "      AY        - Assessment Year to filter by, e.g. 2026-27",
        "      Status    - LEAVE BLANK. The tool fills this automatically.",
        "",
        "3. Delete the grey sample row before running.",
        "4. Save the file, then upload it in the tool and click Run.",
        "",
        "The tool logs in for each row, reads the latest return status from the",
        "Income-Tax portal, and fills it into the Status column. When the run",
        "finishes, click 'Download Updated Excel' to save a copy with the results.",
    ]
    for r, line in enumerate(notes, start=1):
        cell = info.cell(row=r, column=1, value=line)
        if r == 1:
            cell.font = Font(bold=True, size=12, color="1F4E78")
    info.column_dimensions["A"].width = 90

    wb.save(path)


def read_rows(path):
    """Return (workbook, worksheet, col_map, rows). rows is a list of dicts
    with keys name/pan/password/ay and an 'excel_row' index."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    # Map headers (case-insensitive) on row 1.
    col_map = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        if h:
            col_map[str(h).strip().lower()] = c
    for needed in ("pan", "password", "ay"):
        if needed not in col_map:
            raise ValueError(f"Column '{needed.upper()}' not found in the sheet header.")
    if "status" not in col_map:
        col_map["status"] = ws.max_column + 1
        ws.cell(row=1, column=col_map["status"], value="Status")

    rows = []
    for r in range(2, ws.max_row + 1):
        pan = ws.cell(row=r, column=col_map["pan"]).value
        pw = ws.cell(row=r, column=col_map["password"]).value
        if not pan or not pw:
            continue
        pan_s = str(pan).strip()
        # skip the grey sample row from the template
        if pan_s.upper() == "ABCDE1234F":
            continue
        rows.append({
            "excel_row": r,
            "name": str(ws.cell(row=r, column=col_map.get("name", 0)).value or "").strip()
                    if col_map.get("name") else "",
            "pan": pan_s,
            "password": str(pw),
            "ay": str(ws.cell(row=r, column=col_map["ay"]).value or "").strip(),
        })
    return wb, ws, col_map, rows


def save_updated_excel(src_path, results, dst_path):
    """Load the uploaded workbook, write the collected statuses into the Status
    column, and save a NEW copy to dst_path. The original file is left
    untouched. `results` maps excel_row -> status."""
    wb, ws, col_map, _ = read_rows(src_path)
    for excel_row, status in results.items():
        ws.cell(row=excel_row, column=col_map["status"], value=status)
    wb.save(dst_path)


# =========================================================================== #
#  PART 4 — THE BATCH WORKER (runs on a background thread)                      #
# =========================================================================== #
def run_batch(excel_path, headless, log, stop_event, on_result):
    if not _OPENPYXL_OK:
        log("ERROR: openpyxl is not installed. Run:  pip install openpyxl")
        return

    try:
        _, _, _, rows = read_rows(excel_path)
    except Exception as e:
        log(f"ERROR reading Excel: {e}")
        return

    if not rows:
        log("No data rows found. Fill the template (and remove the grey sample row).")
        return

    log(f"Loaded {len(rows)} row(s).")
    log("=" * 60)

    for i, row in enumerate(rows, start=1):
        if stop_event.is_set():
            log("Stopped by user.")
            break

        name, pan, pw, ay = row["name"], row["pan"], row["password"], row["ay"]
        log(f"[{i}/{len(rows)}] {name or pan}  (PAN {pan}, AY {ay or '—'})")

        portal = None
        status = "Error"
        try:
            portal = IncomeTaxPortal(headless=headless, log=log)
            portal.login(pan, pw)

            go_to_filed_returns(portal)
            if ay:
                filter_by_ay(portal, ay)

            status = read_latest_status(portal)
            log(f"  STATUS: {status}")

            logout(portal)

        except InvalidPasswordError:
            status = "Invalid password"
            log("  STATUS: Invalid password — skipping")
        except LoginError as e:
            status = "Login failed"
            log(f"  STATUS: Login failed ({e})")
        except Exception as e:
            status = f"Error: {e}"
            log(f"  ERROR: {e}")
        finally:
            on_result(row["excel_row"], status)
            if portal:
                portal.quit()
            time.sleep(1.0)   # small gap before next login
        log("-" * 60)

    log("=" * 60)
    log("Done. Click 'Download Updated Excel' to save your results.")


# =========================================================================== #
#  PART 5 — TKINTER GUI                                                         #
# =========================================================================== #
class App:
    def __init__(self, root):
        self.root = root
        root.title("Income-Tax Returns — Bulk Status Checker")
        root.geometry("760x560")
        root.minsize(680, 480)

        self.excel_path = tk.StringVar(value="")
        self.headless = tk.BooleanVar(value=False)
        self.results = {}            # excel_row -> status

        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None

        self._build_ui()
        self._poll_log()

        if not _OPENPYXL_OK:
            self._log("WARNING: 'openpyxl' is not installed — the Excel features "
                      "will not work.\n         Install it with:  pip install openpyxl")

    # ---- layout ----------------------------------------------------------- #
    def _build_ui(self):
        pad = dict(padx=10, pady=6)

        head = ttk.Label(self.root, text="Bulk verification-status checker for filed Income-Tax returns",
                         font=("Segoe UI", 12, "bold"))
        head.pack(anchor="w", padx=12, pady=(12, 2))
        ttk.Label(self.root,
                  text="Flow:  download template  ›  fill it  ›  upload  ›  Run  ›  download updated Excel",
                  foreground="#555").pack(anchor="w", padx=12, pady=(0, 8))

        # Step 1 — template
        f1 = ttk.LabelFrame(self.root, text="Step 1 — Get the Excel template")
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="Download Excel Template…",
                   command=self.on_template).pack(side="left", padx=8, pady=8)
        ttk.Label(f1, text="Columns: Name · PAN · Password · AY · Status (auto-filled)",
                  foreground="#555").pack(side="left", padx=6)

        # Step 2 — upload
        f2 = ttk.LabelFrame(self.root, text="Step 2 — Upload your filled template")
        f2.pack(fill="x", **pad)
        ttk.Button(f2, text="Upload Excel…", command=self.on_upload).pack(side="left", padx=8, pady=8)
        ttk.Label(f2, textvariable=self.excel_path, foreground="#1F4E78").pack(side="left", padx=6)

        # Step 3 — run
        f4 = ttk.LabelFrame(self.root, text="Step 3 — Run")
        f4.pack(fill="x", **pad)
        self.run_btn = ttk.Button(f4, text="▶  Run", command=self.on_run)
        self.run_btn.pack(side="left", padx=8, pady=8)
        self.stop_btn = ttk.Button(f4, text="■  Stop", command=self.on_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=4, pady=8)
        ttk.Checkbutton(f4, text="Headless (hide browser — not recommended; OTP/CAPTCHA need the window)",
                        variable=self.headless).pack(side="left", padx=12)

        # Step 4 — download updated Excel
        f5 = ttk.LabelFrame(self.root, text="Step 4 — Get your results")
        f5.pack(fill="x", **pad)
        self.download_btn = ttk.Button(f5, text="Download Updated Excel…",
                                       command=self.on_download_excel, state="disabled")
        self.download_btn.pack(side="left", padx=8, pady=8)
        ttk.Label(f5, text="Saves a NEW copy with the Status column filled in (original untouched).",
                  foreground="#555").pack(side="left", padx=6)

        # Log
        flog = ttk.LabelFrame(self.root, text="Progress log")
        flog.pack(fill="both", expand=True, **pad)
        self.log_box = scrolledtext.ScrolledText(flog, height=12, wrap="word",
                                                 font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, padx=6, pady=6)
        self.log_box.configure(state="disabled")

    # ---- button handlers -------------------------------------------------- #
    def on_template(self):
        if not _OPENPYXL_OK:
            messagebox.showerror("Missing dependency",
                                 "openpyxl is not installed.\n\nRun:  pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(
            title="Save template as…",
            defaultextension=".xlsx",
            initialfile="ITR_status_template.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            export_template(path)
            self._log(f"Template saved: {path}")
            messagebox.showinfo("Template saved",
                                f"Template created:\n{path}\n\nFill it, then upload it in Step 2.")
        except Exception as e:
            messagebox.showerror("Could not save template", str(e))

    def on_upload(self):
        path = filedialog.askopenfilename(
            title="Select your filled template",
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not path:
            return
        self.excel_path.set(path)
        self._log(f"Loaded Excel: {path}")
        # quick preview of row count
        if _OPENPYXL_OK:
            try:
                _, _, _, rows = read_rows(path)
                self._log(f"  {len(rows)} data row(s) detected.")
            except Exception as e:
                self._log(f"  WARNING: {e}")

    def on_download_excel(self):
        if not self.results:
            messagebox.showinfo("Nothing yet",
                                "Run the batch first — there are no statuses to save.")
            return
        path = filedialog.asksaveasfilename(
            title="Save updated Excel as…",
            defaultextension=".xlsx",
            initialfile="ITR_status_updated.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            save_updated_excel(self.excel_path.get(), dict(self.results), path)
            self._log(f"Updated Excel saved: {path}")
            messagebox.showinfo("Saved", f"Updated Excel saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Could not save", str(e))

    def on_run(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.excel_path.get():
            messagebox.showwarning("Missing", "Upload your filled Excel template first (Step 2).")
            return
        if not os.path.isfile(self.excel_path.get()):
            messagebox.showerror("Not found", "The Excel file no longer exists.")
            return

        if not messagebox.askyesno(
                "Start batch?",
                "An Edge window will open and log in for each row in turn.\n\n"
                "Keep the browser visible so you can solve any OTP/CAPTCHA.\n\n"
                "Start now?"):
            return

        self.results = {}
        self.stop_event.clear()
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.download_btn.configure(state="disabled")
        self._log("\n>>> Starting batch…\n")

        self.worker = threading.Thread(
            target=self._worker_wrap,
            args=(self.excel_path.get(), self.headless.get()),
            daemon=True)
        self.worker.start()

    def on_stop(self):
        self.stop_event.set()
        self._log("Stop requested — will halt after the current row…")
        self.stop_btn.configure(state="disabled")

    def _worker_wrap(self, excel_path, headless):
        try:
            run_batch(excel_path, headless,
                      log=lambda m: self.log_queue.put(m),
                      stop_event=self.stop_event,
                      on_result=self._store_result)
        except Exception as e:
            self.log_queue.put(f"FATAL: {e}")
        finally:
            self.log_queue.put("__DONE__")

    def _store_result(self, excel_row, status):
        self.results[excel_row] = status

    # ---- thread-safe logging --------------------------------------------- #
    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _poll_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg == "__DONE__":
                    self.run_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    if self.results:
                        self.download_btn.configure(state="normal")
                else:
                    self._log(msg)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_log)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")   # nicer on Windows; ignored elsewhere
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
