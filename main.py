"""
Lahimena Tours Devis Generation Application

Main entry point for the application
"""

import sys
import tkinter as _tk
from tkinter import messagebox, ttk

import customtkinter as ctk
from customtkinter.windows.ctk_tk import CTk as _CTkClass
from customtkinter.windows.ctk_toplevel import CTkToplevel as _CTkToplevelClass
from customtkinter.windows.widgets.appearance_mode.appearance_mode_tracker import (
    AppearanceModeTracker,
)
from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker

# ── Patch: ignorer silencieusement les erreurs de double-suppression de
# commandes Tcl lors de la destruction de widgets CTK (race condition connue).
_orig_deletecommand = _tk.Misc.deletecommand


def _safe_deletecommand(self, name):
    try:
        _orig_deletecommand(self, name)
    except Exception:
        pass


_tk.Misc.deletecommand = _safe_deletecommand


def _patch_customtkinter_trackers():
    """Clean up CTk tracker loops before windows are destroyed."""
    if getattr(ctk, "_lahimena_tracker_patch_applied", False):
        return

    def _appearance_is_alive(app):
        try:
            return bool(app.winfo_exists())
        except Exception:
            return False

    def _cancel_appearance_after(app=None):
        after_id = getattr(AppearanceModeTracker, "_lahimena_after_id", None)
        after_app = getattr(AppearanceModeTracker, "_lahimena_after_app", None)
        if after_id and after_app is not None and (app is None or after_app is app):
            try:
                after_app.after_cancel(after_id)
            except Exception:
                pass
        if app is None or after_app is app:
            AppearanceModeTracker._lahimena_after_id = None
            AppearanceModeTracker._lahimena_after_app = None

    def _schedule_appearance_update():
        _cancel_appearance_after()
        AppearanceModeTracker.app_list = [
            app for app in AppearanceModeTracker.app_list if _appearance_is_alive(app)
        ]
        for app in AppearanceModeTracker.app_list:
            try:
                AppearanceModeTracker._lahimena_after_app = app
                AppearanceModeTracker._lahimena_after_id = app.after(
                    AppearanceModeTracker.update_loop_interval,
                    AppearanceModeTracker.update,
                )
                AppearanceModeTracker.update_loop_running = True
                return
            except Exception:
                continue
        AppearanceModeTracker.update_loop_running = False

    def _patched_appearance_add(cls, callback, widget=None):
        cls.callback_list.append(callback)
        if widget is not None:
            app = cls.get_tk_root_of_widget(widget)
            if app not in cls.app_list:
                cls.app_list.append(app)
            if not cls.update_loop_running:
                _schedule_appearance_update()

    def _patched_appearance_update(cls):
        if cls.appearance_mode_set_by == "system":
            new_appearance_mode = cls.detect_appearance_mode()
            if new_appearance_mode != cls.appearance_mode:
                cls.appearance_mode = new_appearance_mode
                cls.update_callbacks()
        _schedule_appearance_update()

    def _scaling_is_alive(window):
        try:
            return bool(window.winfo_exists())
        except Exception:
            return False

    def _prune_dead_scaling_windows():
        dead_windows = [
            window
            for window in list(ScalingTracker.window_widgets_dict.keys())
            if not _scaling_is_alive(window)
        ]
        for window in dead_windows:
            ScalingTracker.window_widgets_dict.pop(window, None)
            ScalingTracker.window_dpi_scaling_dict.pop(window, None)

    def _cancel_scaling_after(window=None):
        after_id = getattr(ScalingTracker, "_lahimena_after_id", None)
        after_window = getattr(ScalingTracker, "_lahimena_after_window", None)
        if after_id and after_window is not None and (window is None or after_window is window):
            try:
                after_window.after_cancel(after_id)
            except Exception:
                pass
        if window is None or after_window is window:
            ScalingTracker._lahimena_after_id = None
            ScalingTracker._lahimena_after_window = None

    def _schedule_scaling_check(delay=None):
        _cancel_scaling_after()
        _prune_dead_scaling_windows()
        for window in list(ScalingTracker.window_widgets_dict.keys()):
            try:
                ScalingTracker._lahimena_after_window = window
                ScalingTracker._lahimena_after_id = window.after(
                    ScalingTracker.update_loop_interval if delay is None else delay,
                    ScalingTracker.check_dpi_scaling,
                )
                ScalingTracker.update_loop_running = True
                return
            except Exception:
                continue
        ScalingTracker.update_loop_running = False

    def _patched_scaling_add_widget(cls, widget_callback, widget):
        window_root = cls.get_window_root_of_widget(widget)
        cls.window_widgets_dict.setdefault(window_root, []).append(widget_callback)
        if window_root not in cls.window_dpi_scaling_dict:
            cls.window_dpi_scaling_dict[window_root] = cls.get_window_dpi_scaling(
                window_root
            )
        if not cls.update_loop_running:
            _schedule_scaling_check(delay=100)

    def _patched_scaling_add_window(cls, window_callback, window):
        cls.window_widgets_dict.setdefault(window, []).append(window_callback)
        if window not in cls.window_dpi_scaling_dict:
            cls.window_dpi_scaling_dict[window] = cls.get_window_dpi_scaling(window)
        if not cls.update_loop_running:
            _schedule_scaling_check(delay=100)

    def _patched_scaling_check(cls):
        new_scaling_detected = False
        _prune_dead_scaling_windows()
        for window in list(cls.window_widgets_dict.keys()):
            try:
                if window.winfo_exists() and window.state() != "iconic":
                    current_dpi_scaling_value = cls.get_window_dpi_scaling(window)
                    if current_dpi_scaling_value != cls.window_dpi_scaling_dict[window]:
                        cls.window_dpi_scaling_dict[window] = current_dpi_scaling_value

                        if sys.platform.startswith("win"):
                            window.attributes("-alpha", 0.15)

                        window.block_update_dimensions_event()
                        cls.update_scaling_callbacks_for_window(window)
                        window.unblock_update_dimensions_event()

                        if sys.platform.startswith("win"):
                            window.attributes("-alpha", 1)

                        new_scaling_detected = True
            except Exception:
                continue
        _schedule_scaling_check(
            ScalingTracker.loop_pause_after_new_scaling
            if new_scaling_detected
            else ScalingTracker.update_loop_interval
        )

    def _cleanup_customtkinter_window(window):
        _cancel_appearance_after(window)
        AppearanceModeTracker.app_list = [
            app for app in AppearanceModeTracker.app_list if app is not window
        ]

        _cancel_scaling_after(window)
        ScalingTracker.window_widgets_dict.pop(window, None)
        ScalingTracker.window_dpi_scaling_dict.pop(window, None)

        if AppearanceModeTracker.app_list:
            _schedule_appearance_update()
        else:
            AppearanceModeTracker.update_loop_running = False

        if ScalingTracker.window_widgets_dict:
            _schedule_scaling_check()
        else:
            ScalingTracker.update_loop_running = False

    original_ctk_destroy = _CTkClass.destroy
    original_ctk_toplevel_destroy = _CTkToplevelClass.destroy

    def _patched_ctk_destroy(self):
        _cleanup_customtkinter_window(self)
        original_ctk_destroy(self)

    def _patched_ctk_toplevel_destroy(self):
        _cleanup_customtkinter_window(self)
        original_ctk_toplevel_destroy(self)

    AppearanceModeTracker.add = classmethod(_patched_appearance_add)
    AppearanceModeTracker.update = classmethod(_patched_appearance_update)
    ScalingTracker.add_widget = classmethod(_patched_scaling_add_widget)
    ScalingTracker.add_window = classmethod(_patched_scaling_add_window)
    ScalingTracker.check_dpi_scaling = classmethod(_patched_scaling_check)
    _CTkClass.destroy = _patched_ctk_destroy
    _CTkToplevelClass.destroy = _patched_ctk_toplevel_destroy
    ctk._lahimena_tracker_patch_applied = True


_patch_customtkinter_trackers()

from config import (
    APP_GEOMETRY,
    APP_TITLE,
    APPEARANCE_MODE,
    BUTTON_GREEN,
    BUTTON_GREEN_HOVER,
    BUTTON_RED,
    DEFAULT_COLOR_THEME,
    INPUT_BG_COLOR,
    MAIN_BG_COLOR,
    MUTED_TEXT_COLOR,
    TEXT_COLOR,
)
from gui.main_content import MainContent
from gui.sidebar import Sidebar
from utils.logger import logger


def _launch_main_app(user):
    """Lance la fenêtre principale après authentification."""
    app = ctk.CTk()
    app.title(APP_TITLE)
    app.geometry(APP_GEOMETRY)
    app.configure(fg_color=MAIN_BG_COLOR)

    def _apply_ui_theme_inner():
        cursor_color = TEXT_COLOR
        app.option_add("*Entry.insertBackground", cursor_color)
        app.option_add("*Text.insertBackground", cursor_color)
        app.option_add("*Entry.selectBackground", BUTTON_GREEN)
        app.option_add("*Text.selectBackground", BUTTON_GREEN)

        style = ttk.Style(app)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        bg_main = MAIN_BG_COLOR
        bg_input = INPUT_BG_COLOR
        fg_main = TEXT_COLOR
        fg_muted = MUTED_TEXT_COLOR
        accent = BUTTON_GREEN
        accent_hover = BUTTON_RED
        border = "#C9DDE3"

        style.configure(".", background=bg_main, foreground=fg_main, fieldbackground=bg_input)
        style.configure("TCombobox", foreground=fg_main, fieldbackground=bg_input,
                        background=bg_input, bordercolor=border, lightcolor=border,
                        darkcolor=border, arrowcolor=fg_muted)
        style.map("TCombobox", fieldbackground=[("readonly", bg_input)],
                  foreground=[("readonly", fg_main)])
        style.configure("Treeview", background=bg_input, fieldbackground=bg_input,
                        foreground=fg_main, bordercolor=border, rowheight=26)
        style.map("Treeview", background=[("selected", accent)],
                  foreground=[("selected", "#FFFFFF")])
        style.configure("Treeview.Heading", background=bg_main, foreground=fg_main,
                        bordercolor=border, relief="flat", font=("Poppins", 10, "bold"))
        style.map("Treeview.Heading", background=[("active", accent_hover)],
                  foreground=[("active", "#FFFFFF")])
        style.configure("Vertical.TScrollbar", background=bg_input, troughcolor=bg_main,
                        bordercolor=border, arrowcolor=fg_muted)
        style.configure("Horizontal.TScrollbar", background=bg_input, troughcolor=bg_main,
                        bordercolor=border, arrowcolor=fg_muted)

    _apply_ui_theme_inner()

    app.grid_columnconfigure(0, weight=0)
    app.grid_columnconfigure(1, weight=1)
    app.grid_rowconfigure(0, weight=1)

    main_content = MainContent(app)
    _sidebar = Sidebar(app, main_content.update_content)

    logger.info(f"Application démarrée — utilisateur : {user['username']} ({user['role']})")
    app.mainloop()


def main():
    """Main application entry point"""
    try:
        logger.info(f"Starting {APP_TITLE}")

        # Set appearance theme
        ctk.set_appearance_mode(APPEARANCE_MODE)
        ctk.set_default_color_theme(DEFAULT_COLOR_THEME)
        logger.debug(f"Theme set: {APPEARANCE_MODE} mode, {DEFAULT_COLOR_THEME} color")

        # ── Écran de login ────────────────────────────────────────────────
        from gui.forms.login_form import LoginWindow

        logger.info("Application started successfully")
        login = LoginWindow(on_login_success=_launch_main_app)
        login.mainloop()
        # _launch_main_app lance sa propre boucle après le login

    except Exception as e:
        error_msg = f"Application error: {e}"
        logger.error(error_msg, exc_info=True)
        # Show error dialog to user
        try:
            messagebox.showerror(
                "❌ Erreur Application",
                f"Une erreur est survenue:\n\n{str(e)}\n\nVoir les logs pour plus de détails.",
            )
        except Exception as dialog_error:
            print(f"CRITICAL ERROR: {error_msg}")
            print(f"Dialog error: {dialog_error}")
        raise


if __name__ == "__main__":
    main()
