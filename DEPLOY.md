# 考勤工作台部署

## Cloudflare Workers 与静态资源

在项目目录运行 `npm ci` 安装依赖。使用 `npm run dev` 预览，使用 `npx wrangler login` 登录后运行 `npm run deploy` 发布。

Wrangler 从 `templates/` 发布静态页面，`worker/index.ts` 处理同域的 `POST /process`。请求中的 Excel 文件只在生成报表时处理。发布时无需上传员工表格或配置存储。

考勤日历集中在 `worker/calendar.ts`。新增年份时，按国务院办公厅通知录入放假起止日期、补班日期和来源链接，再发布 Worker。未配置的年份按普通周历处理，并在导出的「考勤日历」中标明。

`wrangler.jsonc` 已声明正式域名 `saidi.clickadwave.org`，发布时 Wrangler 会绑定此域名。要单独创建预览环境，可使用其他 Worker 名称和不包含正式域名的 Wrangler 配置。保留旧 Caddy 与容器，便于需要时回退。

## Docker 与 Caddy

域名为 `saidi.clickadwave.org`。服务器上由现有 Caddy 接收 HTTPS 请求，再转发给独立的 Docker Compose 项目 `saidi-attendance`。

## Caddy 运行在宿主机

1. 将 `app.py`、`templates/`、`requirements.txt`、`Dockerfile`、`docker-compose.yml` 和 `.dockerignore` 上传到服务器上的独立目录。员工表格通过网页上传。
2. 确认服务器的 `18081` 端口空闲。在项目目录执行：

   ```sh
   docker compose up -d --build
   docker compose ps
   curl -f http://127.0.0.1:18081/ -o /dev/null
   ```

   如果端口已被占用，在项目的 `.env` 设置 `SAIDI_PORT=其他空闲端口`，并将 Caddy 上游端口改为相同值。
3. 把 `Caddyfile.saidi` 中的站点块合并到现有 Caddyfile。如果已有同名站点，检查其当前用途后再修改该站点块。以系统服务和默认配置路径为例，先校验，再平滑加载：

   ```sh
   sudo caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile
   sudo systemctl reload caddy
   curl -f https://saidi.clickadwave.org/ -o /dev/null
   ```

公网入口沿用现有 Caddy 的 80 和 443 端口，应用端口只绑定服务器的回环地址。Caddy 会在满足域名验证条件时自动签发和续期证书。若服务器已有认证网关，可将该子域名接入现有访问控制。也可在此站点块配置 Caddy 的 `basic_auth`，密码哈希用 `caddy hash-password` 生成。

## Caddy 运行在 Docker 容器内

让考勤服务加入现有 Caddy 使用的 Docker 网络。在考勤项目下创建 `docker-compose.override.yml`，将示例中的 `caddy_network` 替换为现有网络名：

```yaml
services:
  attendance:
    networks:
      caddy:
        aliases:
          - saidi-attendance-app
networks:
  caddy:
    external: true
    name: caddy_network
```

此时将子域名站点块中的上游改为 `reverse_proxy saidi-attendance-app:8080`。在现有 Caddy 容器内，用其实际配置路径执行 `caddy validate` 和 `caddy reload`。回环端口仍可用于宿主机排查。

回退时，从 Caddy 配置移除本次新增的站点块并重新加载，然后在考勤项目目录运行 `docker compose down`。
