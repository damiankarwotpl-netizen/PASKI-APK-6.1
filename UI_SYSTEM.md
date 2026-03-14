# UI System Notes

This repository now includes a lightweight in-code design system for the Kivy app:

- Shared interactive controls (`ModernButton`, `ModernInput`, `ColorSafeLabel`).
- Reusable screen shell API in `FutureApp`:
  - `build_screen_shell(title, subtitle, back_to)`
  - `styled_popup(title, content, size_hint)`

## Goals

- Keep visual consistency across modules.
- Make placeholder modules production-ready in layout and UX.
- Reduce duplicated UI boilerplate in future features.

## Integration points

Implemented in `main.py` and used by:
- Template screen
- Contacts screen and contact cards
- Global message popup and logs popup
- Cars / Paski / Pracownicy / Zakłady / Settings screens

## Future extension

- Migrate remaining ad-hoc popups to `styled_popup`.
- Add iconography and semantic color scale (success/warning/error).
- Split UI components into separate module when app grows further.
