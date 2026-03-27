# Mute

Minimal Windows app (wxPython) that toggles mute for the foreground app using a hotkey.

## Run

1. Create and activate a virtual environment.
2. Install dependencies:

```
pip install -r requirements.txt
```

3. Run the app:

```
python src/mute/main.py
```

## Hotkey format

- Default: `pause` (the Pause/Break key).
- Click "Record" and press a key or mouse button (mouse4/mouse5 supported).
- You can also use pynput style strings like `<ctrl>+<alt>+m`.

## Notes

- The toggle affects the foreground app's audio session when one exists.
- "Run on Windows startup" uses the current user Run registry key.
- "Minimize to tray on close" keeps the app running in the tray.
