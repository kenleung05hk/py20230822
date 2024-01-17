from translate import Translator
translator= Translator(to_lang="en",from_lang ="zh")
translation = translator.translate("一隻狗")
print(translation)