#Python substitute with numbering inside patterns

import re
text = """

This is a testing document for testing purposes only.[^0] This is a testing document for testing purposes only. This is a testing document for testing purposes only.[^121][^5] This is a testing document for testing purposes only.

[^0]: Footnote contents.

[^0]: Footnote contents.

[^0]: Footnote contents.
"""

footnote_rex = re.compile(r'\[\^0\]\:')
intext_rex = re.compile(r'(\[\^\d+\](?!\:))')

for idx,  entry in enumerate(footnote_rex.findall(text)):
    print(entry)
    text = text.replace(entry, '[^{}]:'.format(idx+1))

print(text)

for idx, entry in enumerate(intext_rex.findall(text)):
    text = text.replace(entry, '[^{}]'.format(idx+1))



print(text)