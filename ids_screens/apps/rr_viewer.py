from pywinauto.application import Application


def shoot_rr_viewer(conf):
    app = Application(backend="win32").start(conf["path"])
    app.wait_cpu_usage_lower()

    win = app.window(class_name="TMainForm")

    win.move_window(200, 200, 600, 400)
    app.wait_cpu_usage_lower()

    # ROSRaveReportsViewer.bmp
    win.capture_as_image().save("ROSRaveReportsViewer.bmp")

    win.close()
