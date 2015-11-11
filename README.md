# SPTools
The purpose of this extension is to help **developers** to work on remote **SharePoint** sites using **Visual Studio Code**.

## Install
------

`ext install SPTools`

## config.json
------

This is your extension **global** config file.

```
{
	"path": "C:\\work\\",
	"folders": [
		"/_catalogs/masterpage/sptools",
		"/style library/en-us/themable/tools"
	]
}
```

### "path" (String)

Specify your top level local workspace folder here. This should typically be updated after installation.

This is where the new SP workspaces will be created, using the project name as subfolder name.

**Example**

Using `"path": "C:\\work\\"` will create here a folder per new workspace.

### "folders" (Array)

Specify the remote folders you want to download and sync in your workspace.

The folder structures will be replicated locally from the workspace root.

**Example**

Using the `SPTools: Init` command, with `test` as project name and the initial configuration above will create:

- `C:\\work\\test\\`
- `C:\\work\\test\\_catalogs\\masterpage\\sptools`
- `C:\\work\\test\\style library\\en-us\\themable\\sptools`

## spconfig.json
------

This is a **workspace specific** config file, created automatically when using the `SPTools: Init` command at the root of the workspace.

For now it only stores the SharePoint site URL.

## Commands
------

### SPTools: Init

Initialize and sync a SharePoint workspace.

Replicate folder structure and download all files locally. 

### SPTools: Check file freshness

Compare remote and local last modified dates and update statusbar indicators.

### SPTools: Sync file

Sync current file

### SPTools: Upload file

Upload current file

### SPTools: Check in file

Check in current file

### SPTools: Check out file

Check out current file

### SPTools: Discard file check out

Discard current file check out (use with caution)

### SPTools: Sync entire workspace

Sync entire workspace (use with caution)

### SPTools: Reset credentials cache

Delete all saved credentials (use with caution)

## A word about credentials
------

Credentials are not stored in files, but rather kept in the extension memory.

Use `SPTools: Reset credentials cache` command to reset the cache.

** Enjoy!**