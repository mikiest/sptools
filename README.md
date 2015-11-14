# SPTools
The purpose of this extension is to help **developers** to work on remote **SharePoint** sites using **Visual Studio Code**.

## Install
------

`ext install SPTools`

## Extension settings
------

These are your extension settings:

```json
"sptools.workFolder": {
	"type": "string",
	"default": "c:\\\\work\\\\",
	"description": "Path to your top level work folder"
},
"sptools.spFolders": {
	"type": "array",
	"default": [
		"/_catalogs/masterpage/sptools",
		"/style library/en-us/themable/sptools"
	],
	"description": "Folders to be fetched from SharePoint sites"
},
"sptools.storeCredentials": {
	"type": "boolean",
	"default": true,
	"description": "Set to false if you don't want credentials to be cached"
}
```

To override them, specify your own values in `File > Preferences` then either `User Settings` or `Workspace Settings`.

:bulb: It's a good idea to set your own SharePoint folders list, the default ones will probably not exist on your sites.

### Folders creation

Please keep in mind that any non existing folder used in the config will be created.

## spconfig.json
------

This is a **workspace specific** config file, created automatically at the root of the workspace when using the `SPTools: Init` command.

For now it only stores the SharePoint site URL.

## Commands
------

### SPTools: New workspace

Create a new SharePoint workspace from a remote site.

Replicate folder structure and download all files locally.

SharePoint workspaces (folders with a top level spconfig.json file) will prompt you to login to the relevant remote site the first time you open a file. This is because the extension is checking the current file status (date and check out status) when opening it (there is a small latence so it won't check dozen of files if you quickly switch between them).

### SPTools: Check file freshness

**Compare** remote and local last modified dates and update statusbar indicators accordingly.

### SPTools: Sync file

**Download** remote file, replace current local file.

### SPTools: Upload file

**Upload** current file to remote site, replace remote file with local one.

### SPTools: Check in/out file

Check file status and ask if you want to **check it in, out or discard** the current check out depending on the current state.

### SPTools: Sync entire workspace

Sync entire workspace (use with caution)

### SPTools: Reset credentials cache

**Delete all cached credentials**

## A word about credentials
------

Credentials are not stored in files, but rather kept in the extension memory.

Use `SPTools: Reset credentials cache` command to reset the cache.

** Enjoy!**