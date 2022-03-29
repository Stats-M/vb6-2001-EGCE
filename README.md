# vb6-2001-EGCE (elite-games conference extractor)
Rade (C) 2001 (MS Visual Basic 6) Learning project of my youth. Hail to spaghetti code and a lot of emotional comments!
Specialized old BBS-style web conference data extractor

This project was commited to one of the famous russian webcites dedicated to all space-related games - Elite Games (currently located at www.elite-games-.ru) It was (and is) the central place for fans of Elite, Privateer, EVE, X-series and other games in the genre of space simulators.

Back that time (2000-2001) this project was still a fan-cite, located at a free web-hosting, running a simple bulletin-board system. You can recall topics in threads of these BBs looked like a trees with indent. Somewhere at 2001 the cote finally has moved to a paid hosting and implements modern PHP forum script.

In order to preserve the fun and feel of the old forum, the happines of the good old days, Elite Game Conference Extractor (EGCE) project has been started.
The task were:
-download all text content
-rebuild all it in a manner of a modern PHP forum
-implement navigation system for all
-preserve old BBs tree-like thread index because indentation in BBs indicates to which post a user has publish a reply.

So EGCE does:
-download all BBs content
-generate forum pages, table of content and java scripts (sorting, indentation, etc.)
NB. In modern browser identations script works no more, keeping thread trees "flat"
