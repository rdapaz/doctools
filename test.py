        import re
        text = 'how are u? umberella u! u. U. U@ U# u '
        rex = re.compile(r'\bu\b', re.IGNORECASE)
        print (rex.sub ('you', text))