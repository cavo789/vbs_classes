# Windows.vbs

This script exposes a VB Script class for working with the operating system.

## Table of content

- [getPreferredLanguage](#getpreferredlanguage)
	- [Sample script](#sample-script)

## getPreferredLanguage

Read from the registry the user's preferred languages and return the first mentionned between dutch or french : in, f.i., Internet Explorer, the user can mention that he can read english text, dutch, french, spannish, ...

By adding multiple supported languages, he'll specify his preferences : first french, then english, then ...

This function will return the first "preferred" language between French or Dutch.

Note : this function will only return FR (for French) or NL (for Dutch), if needed, just modify the function.

### Sample script

```vbnet
Set cWindows = New clsWindows
wScript.echo "You prefer " & cWindows.getPreferredLanguage("FR") & _
	" content"
Set cWindows = Nothing
```