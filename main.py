import csvparser
import GUI

csvparser = csvparser.csvParser()

gui = GUI.GUIWindow()
gui.startGUI(csvparser)