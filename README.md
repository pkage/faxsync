# FaxSync

Windows faxes are not stored in the filesystem like [normal systems](https://en.wikipedia.org/wiki/Linux) do, so you have to do an entire song and dance to get all faxes on a machine.

This should build a `FaxSync.exe` executable, which can then be used to either extract the last N days of faxes (default 1 week):

```
> FaxSync.exe \\myfaxserver 10
```

Or you can show some information about the fax server

```
> FaxSync.exe \\myfaxserver info
```

## Creating a windows fax service api client in VB

Don't do this.

I cannot stress that enough--do not do this. It's a soul crushing journey through incomplete documentation, old Windows APIs, and boneheaded design decision.

If you do, here are some notes:

 - Faxes on windows are handled by an archaic API. Inbound faxes are hidden on the filesystem, locked away in a proprietary database.
 - Access is provided through the `FAXCOMEXLib` .COM library, which you can use through .NET systems
 - An example is provided from the MS docs (see second link below). This is helpful, but:
    - Note that to use this, you must enable the `FAXCOMEXLib` .COM reference. See Project Explorer > References > Add Reference
    - The specific reference is "Microsoft Fax Service Extended COM Type Library", version 1.0. Something tells me there won't be a version 2.0.
    - You'll also need to add a `Imports FAXCOMEXLib` to the top of the file.
    - Skip the line which has the `objFaxServer.Folders.IncomingArchive.Refresh()` method, this will fail with `0x80070032 Operation Failed.`
       - Googling this will not help. Why is it there? Why is it like this? A mystery lost to the ages.
 - `FaxIncomingMessage`'s `TransmissionEnd` date does not exist. It's an invalid date.
 - (At least with our fax system), `FaxIncomingMessage`'s `SenderFaxNumber` doesn't exist, rather it's a part of the `CallerId` string 

Useful links:
 - [Fax Server Documentation](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/fax/-mfax-faxserve)
 - [Open a fax from the Incoming Fax Archive](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/fax/-mfax-opening-a-fax-from-the-incoming-archive)
