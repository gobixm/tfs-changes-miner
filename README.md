# tfs-changes-miner
to mine changed files with related work items from team foundation server

## usage:

edit settings.json to your requirements

```
{
    "url": "http://tfs-server:port/tfs/collection",
    "target": "$/path to target branch",
    "branches": [{
        "path": "$/path to analyzed branch",
        "from": "changes number to start with",
        "to": "changes number to end with"
    }],
    "ignore": [
        "/Path you want to ignore"
    ],
    "mappings": {
        "/path to folder": "module name(alias)",
    }
}
```
