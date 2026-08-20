
import uno
def run():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
    try:
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    except Exception as e:
        print("Could not connect to LO:", e)
        return
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    if not doc:
        print("No doc open")
        return
    styles = doc.StyleFamilies.getByName("PageStyles")
    for name in styles.getElementNames():
        style = styles.getByName(name)
        print(f"Style: {name}, Height: {style.Height}, TopM: {style.TopMargin}, BotM: {style.BottomMargin}")
        print(f"  HeaderIsOn: {style.HeaderIsOn}, HeaderHeight: {style.HeaderHeight}, HeaderBodyDist: {style.HeaderBodyDistance}")
        print(f"  FooterIsOn: {style.FooterIsOn}, FooterHeight: {style.FooterHeight}, FooterBodyDist: {style.FooterBodyDistance}")
        break
run()
