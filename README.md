# bainultra-n8n

Hosts the configurator workflow at `data/configurator.json`.

## Infrastructure

Deployed via [aliajs](https://github.com/jdecaron/aliajs) in 14 lines of code:

https://github.com/jdecaron/aliajs/blob/erpnext/configurations/instances.js#L90-L103

`n8n/docker-compose.yml` is used by aliajs.

## Run n8n Locally

```
cd n8n
docker-compose up
```

See the full setup instructions in aliajs:
https://github.com/jdecaron/aliajs/blob/erpnext/configurations/instances.js#L90-L103
