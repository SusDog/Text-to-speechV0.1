dim msg, sapi
msg=inputbox("Type your text here.","Text to speech by susdev")
set sapi=createobject("sapi.spvoice")
sapi.speak.msg