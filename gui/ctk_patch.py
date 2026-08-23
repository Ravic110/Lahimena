"""Correctifs appliques a CustomTkinter au demarrage.

CustomTkinter garde deux boucles de surveillance -- apparence et mise a
l'echelle DPI -- qui replanifient un `after` sur des fenetres deja detruites,
et supprime deux fois certaines commandes Tcl a la destruction des widgets.
Ces correctifs remplacent les methodes concernees pour nettoyer avant
destruction.

Ils atteignent onze elements d'API privee, par des chemins de modules profonds.
Tant qu'ils vivaient dans main.py sous forme d'imports nus, une version de
CustomTkinter deplacant l'un de ces modules empechait l'application de demarrer,
et l'echec survenait avant que le gestionnaire d'erreur de main() n'existe :
l'utilisateur voyait une trace Python, pas un message.

Ils sont donc facultatifs. Absents, l'application demarre sans eux ; elle
laissera filer quelques rappels `after` a la fermeture des fenetres, ce qui est
sans commune mesure avec un refus de demarrer. La borne haute posee sur
customtkinter dans pyproject.toml reste la vraie protection ; ceci en est le
filet.
"""

import sys
import tkinter as _tk

import customtkinter as ctk

from utils.logger import logger


def _charger_les_traceurs():
    """Les deux traceurs internes de CustomTkinter et les classes de fenetre.

    Isole pour que les tests puissent simuler un interne deplace.
    """
    from customtkinter.windows.ctk_tk import CTk as _CTkClass
    from customtkinter.windows.ctk_toplevel import CTkToplevel as _CTkToplevelClass
    from customtkinter.windows.widgets.appearance_mode.appearance_mode_tracker import (
        AppearanceModeTracker,
    )
    from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker

    return AppearanceModeTracker, ScalingTracker, _CTkClass, _CTkToplevelClass


def appliquer() -> bool:
    """Installe les correctifs. Renvoie False si CustomTkinter a change.

    Ne leve jamais : demarrer sans correctif vaut mieux que ne pas demarrer.
    """
    try:
        (
            AppearanceModeTracker,
            ScalingTracker,
            _CTkClass,
            _CTkToplevelClass,
        ) = _charger_les_traceurs()
    except Exception as exc:
        logger.warning(
            "Correctifs CustomTkinter non appliques (interne deplace ou absent) : "
            f"{exc}. L'application demarre sans eux."
        )
        return False

    try:
        _installer(AppearanceModeTracker, ScalingTracker, _CTkClass, _CTkToplevelClass)
    except Exception as exc:
        logger.warning(
            f"Correctifs CustomTkinter non appliques : {exc}. "
            "L'application demarre sans eux."
        )
        return False

    return True


def _installer(
    AppearanceModeTracker, ScalingTracker, _CTkClass, _CTkToplevelClass
) -> None:
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
                app
                for app in AppearanceModeTracker.app_list
                if _appearance_is_alive(app)
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
            if (
                after_id
                and after_window is not None
                and (window is None or after_window is window)
            ):
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
                        if (
                            current_dpi_scaling_value
                            != cls.window_dpi_scaling_dict[window]
                        ):
                            cls.window_dpi_scaling_dict[window] = (
                                current_dpi_scaling_value
                            )

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
