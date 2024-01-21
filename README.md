# pyolk
Python parser for Outlook OLK binary caches

## Background
I discovered that a long-deleted email account when I used Outlook had left behind a cache of emails and calendar invites on my Mac - I was curious, so I started chipping away at the binary files to see if I could fully decipher them (enough of the data is stored in unicode or plain text that I could tell what the files were).

I was mostly successful, and think I ended up able to decode and map more than half the fields that Outlook included in the cache, and definitely all of the important ones.

## Usage
By default, on macOS Outlook will have put its cache in `~/Users/harry/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile` - if you just import pyolk and initialize the main class, it'll look in that directory and parse the whole cache. Then calling `export` will create a folder structure mirroring your cached Outlook inbox, and write all of the emails, events, notes, etc. into those folders.

By default `export` writes everything to a new `Recovered Outlook Data` folder in the current working directory.

```
from pyolk import PyOLKReader
p = PyOLKReader()
p.export()
```
