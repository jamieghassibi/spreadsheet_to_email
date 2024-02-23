# TODO Build GUI with kivy

import emailer
import gui

def main():
    gui.main()
    emailer.main(
        gui.files['subject'],
        gui.files['email'],
        gui.files['css'],
        gui.files['csv'],
        gui.files['attachment'],
        gui.files['test'],
        )

main()
