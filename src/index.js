import { Hono } from 'hono'
import { proxy } from 'hono/proxy'

const app = new Hono()

app.all('*', (c) => {
  return proxy(`http://127.0.0.1:5678${c.req.path}`)
})

export default app
