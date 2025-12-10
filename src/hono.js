import { Hono } from 'hono'
import { proxy } from 'hono/proxy'

const app = new Hono()

app.all('/404-uc3C6', (c) => {
  return c.notFound()
})

app.all('*', (c) => {
  return proxy(`http://127.0.0.1:5678${c.req.path}`)
})

export default app
