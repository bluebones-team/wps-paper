function onload(ribbonUI) {
    wps.ribbonUI ??= ribbonUI;
    wps.Enum ??= {
        msoCTPDockPositionLeft: 0,
        msoCTPDockPositionRight: 2
    };
    return !0;
}

//调整格式
function adjustFormat() {
    //字号至少24号
    pre().Slides.loop(Slide => {
        Slide.Shapes.loop(Shape => {
            const font = Shape.TextFrame?.TextRange?.Font;
            if (font?.Size < 24) {
                font.Size = 24;
            }
        })
    });
    return !0;
}

function showHelp() {
    const helpID = wps.PluginStorage.getItem("helpPanel_id");
    if (!helpID) {
        const helpPanel = wps.CreateTaskPane(location.origin + "/ui/help.html");
        wps.PluginStorage.setItem("helpPanel_id", helpPanel.ID);
        helpPanel.Visible = true
    } else {
        const helpPanel = wps.GetTaskPane(helpID)
        helpPanel.Visible = !helpPanel.Visible
    }
    return !0
}