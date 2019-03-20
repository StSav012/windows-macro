# Windows Macro

Emulate mouse and keyboard events to perform routine tasks. The sequence of the actions is written with a specific syntax.

## Disclaimer

This is a Visual Basic 6 project found in my archive. I have never thought that I come back to it.

## Usage Principles

The only window of the application has 4 buttons, a text field and a line that shows the cursor coordinates.

The first two buttons start and stop the execution of the sequence. The other two buttons allow one to insert a code and save the content of the text area.

The supported commands are the following:

* `beep [frequency duration]`, `sound [frequency duration]` produce a system default beep or a defined tone;

* `mv x y`, `move x y`, `moveto x y` move the cursor to the position (`x`, `y`);

* `sleep duration`, `wait duration`, `delay duration` pause the script execution for the given time in *ms*;

* `click [{1 | 2 | 3}]` emulates a mouse button click:

    * `click` and `click 1` emulate the left mouse button click,
    * `click 2` emulates the right mouse button click,
    * `click 3` emulates the middle mouse button click;

* `echo text`, `type text` emulate the text input by keyboard;

* `press {keyname | keycode}`, `presskey {keyname | keycode}` emulate the press of a key denoted by its code or one of the following words:

    * `backspace`, `bksp`,
    * `tab`,
    * `enter`, `return`,
    * `shift`,
    * `ctrl`, `control`,
    * `alt`, `menu`,
    * `esc`, `escape`,
    * `space`,
    * `page up`, `pgup`,
    * `page down`, `pgdn`,
    * `end`,
    * `home`,
    * `left`, `left arrow`, `arrow left`,
    * `up`, `up arrow`, `arrow up`,
    * `right`, `right arrow`, `arrow right`,
    * `down`, `down arrow`, `arrow down`,
    * `print screen`, `prt scr`, `prt sc`, `snapshot`, `screenshot`, `sshot`, `shot`,
    * `delete`, `del`;

* `keydown {keyname | keycode}`, `keydn {keyname | keycode}` emulate only the pressing of a key, not its release;

* `keyup {keyname | keycode}` explicitly releases a key;

* `alert message` and `info message` produce different types of message boxes;

* `launch application`, `start application`, `run application`, `cmd application` try to execute an application;

* `origin [title]` sets the zero point for the cursor coordinates to the upper left corner of a window that has the given title or to the corner of the screen if no title is given or no corresponding window is found;

* `activate title` brings the target window to front;

* `end` finishes the script execution;

* `exit` closes the application *without saving anything*;

* `loop [start=1] end [step=1]` indicates the start of a loop, `next` marks its end, `%%%` stands for the iteration index.

Use `Ctrl + Enter` key sequence to insert the cursor coordinates displayed at the bottom line of the window into the text area. They are used as the `move` command parameters.

All the supported commands are case-insensitive.

The script is paused while `Scroll Lock` in *On*.
