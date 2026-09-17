# 发票报销工具

把要报销的票据丢进一个文件夹，程序自动完成
**解析 → 建台账 → 查重 → 生成报销单（PDF + Excel）→ 合并打印**。

带 Web 界面：本机双击就能用；也可以用 Docker 部署到 NAS，全公司几个人同时在线用，各人有账号、有权限、有审计记录。

支持票据：增值税电子发票、数电票、机票行程单、铁路电子客票（含退票 / 改签红字）、打车行程单、登机牌、OFD / XML。

---

## 目录

- [功能特性](#功能特性)
- [快速开始](#快速开始)
  - [方式一：Windows 本机](#方式一windows-本机)
  - [方式二：Docker / NAS](#方式二docker--nas)
- [配置项](#配置项)
- [数据在哪、怎么备份](#数据在哪怎么备份)
- [角色与权限](#角色与权限)
- [升级到新版本](#升级到新版本)
- [常见问题](#常见问题)
- [开发与测试](#开发与测试)

---

## 功能特性

- **票据解析**：增值税电子发票、数电票、机票行程单、铁路电子客票、打车行程单、登机牌、OFD / XML
- **票面红字**：能识别铁路票上退票 / 改签那行红字，退票费、改签差价不会被当成车票钱
- **出行人识别**：票面 → 手机号对照表 → 文件夹名 三层兜底，另有人工核对面板；出行人和报销人分两列
- **发票与行程单配对**：同笔业务的票挨着输出，附件不重复计钱
- **三种报销单版式**：费用报销单、差旅费报销单、模板二（差旅单支持按人填补助）
- **台账真身是 SQLite**（`发票台账.db`），Excel 只是给人看的导出快照；几个人同时入库、改台账不会互相覆盖
- **多用户**：管理员 / 财务审核员 / 业务人员 / 只读，权限卡口在后端，带审计日志
- **不上传票据也能用本机文件夹**：在 NAS 版界面里可以直接把本机文件夹传上去（保留目录结构）

---

## 快速开始

### 方式一：Windows 本机

1. 装 Python 3.11+（勾选 "Add python.exe to PATH"），然后装依赖：
   ```bat
   pip install openpyxl pypdf pdfminer.six
   ```
2. 双击 **`启动发票报销工具.bat`**，浏览器会自动打开界面。
3. 不设任何 `FB_*` 环境变量时自动是「本机模式」：只监听 `127.0.0.1`、不用登录、5 分钟无操作自动退出，数据就放在程序目录里。

### 方式二：Docker / NAS

**镜像已经构建好了**，NAS 上不用编译，拉下来就能跑。

镜像地址：

```
ghcr.io/zjm117/fapiao-baoxiao:latest
```

#### 1. 准备目录

```bash
mkdir -p /volume1/docker/fapiao
cd /volume1/docker/fapiao
```

（`volume1` 按自己的存储池改，在 NAS 的「文件管理」里能看到实际路径。）

#### 2. 写 `docker-compose.yml`

这个文件是**唯一**需要准备的东西 —— 代码都在镜像里，不用把项目拷到 NAS。

```yaml
services:
  fapiao:
    image: ghcr.io/zjm117/fapiao-baoxiao:latest
    container_name: fapiao-baoxiao
    restart: unless-stopped
    ports:
      # 左边是 NAS 上的端口，8766 被占用就改成别的（比如 18766:8766）
      - "8766:8766"
    environment:
      FB_SERVER: "1"
      FB_HOST: "0.0.0.0"
      FB_PORT: "8766"
      FB_DATA_DIR: "/data"
      FB_SOURCE: "/票据"
      # ↓↓↓ 改成你自己的口令：管理员初始密码 + 无账号时的登录门槛
      FB_PASSWORD: "change-me"
      TZ: "Asia/Shanghai"
    volumes:
      # 台账 / 配置 / 输出 / 备份 / 缓存都在这里，别删
      - ./data:/data
      # 要报销的票据放这里（PDF / OFD / XML / 图片都行）
      - ./票据:/票据
```

> **口令那行要不要写死？** 上面 `FB_PASSWORD: "change-me"` 是直接写字面值，最简单，推荐。
> 如果你的 compose 里是被改成 `FB_PASSWORD: "${FB_PASSWORD}"` 这种写法（NAS 图形界面「新建项目」
> 常常这么生成），那它就是个**占位符**，值要去**同一个目录下的 `.env` 文件**里找：
>
> ```bash
> cd /volume1/docker/fapiao
> printf 'FB_PASSWORD=你的新口令\n' > .env    # 必须是这个目录、这个文件名
> docker compose up -d                          # 环境变量只在启动时读，要重建容器
> ```
>
> ⚠️ **`.env` 里没写这个变量时，compose 会把它当空字符串** —— 而本程序「口令为空 = 不要登录，
> 谁打开都是管理员」。部署完先确认一下：
>
> ```bash
> docker compose config | grep -i FB_PASSWORD   # 看解析出来的实际值，别是空的
> ```

#### 3. 登录镜像仓库（私有镜像只需一次）

```bash
docker login ghcr.io
# 用户名：ZJM117
# 密  码：一个勾了 read:packages 权限的 GitHub Token
#        （GitHub → Settings → Developer settings → Tokens (classic) → Generate new token）
```

> 嫌麻烦可以把 Package 设成 Public 就不用登录，但**镜像里含源码**，设公开等于公开代码，不建议。

#### 4. 起服务

```bash
docker compose pull
docker compose up -d
docker compose logs -f        # 看到「[就绪] 请用浏览器访问」就是好了
```

#### 5. 打开界面

浏览器访问 `http://NAS的IP:8766`

**首次登录**：用户名 `admin`，密码就是你设的 `FB_PASSWORD`。

进去以后：

1. 在「**设置 → 账号与权限**」给同事开各自的账号（每人选一个角色），他们就不必知道这个口令了
2. 「**凭证入库**」页 → 挑票据在哪：
   - **上传本机文件夹**：点「上传本机文件夹…」或直接把文件夹拖进页面，选的是**你这台电脑**的目录，
     传上去落在 NAS 的 `票据/上传/<文件夹名>/`，**目录结构原样保留**（发票和它的行程单必须待在一起，配对才认得出来）
   - **用 NAS 上已有的目录**：路径框填容器内路径，比如 `/票据`
3. 「**凭证台账**」页 → 查询、勾选、点「出行人核对」把出行人过一遍
4. 「**报销单**」页 → 选单据类型 → 生成 PDF + Excel
5. 生成的单据在「**输出**」列表里，点**打开**是浏览器里直接看 PDF，点**定位**是下载

#### 用 NAS 的图形界面部署（绿联 / 群晖 / 威联通）

绿联 UGOS Pro：「Docker」→「项目」→「新建项目」→ 项目名填 `fapiao`、路径选到刚才那个目录 → 把上面的
`docker-compose.yml` 内容粘进去 → 点部署。群晖 / 威联通只是入口位置不同，compose 内容一模一样。

#### 可选：自己构建镜像

不用现成镜像的话，把仓库 clone 到 NAS，用仓库里的 `docker-compose.yml`（放开 `build: .` 那行）：

```bash
git clone https://github.com/ZJM117/fapiao-baoxiao.git
cd fapiao-baoxiao
docker compose up -d --build     # 第一次要几分钟：下 chromium + 中文字体
```

也可以在本机构建好推到你自己的镜像仓库：

```bash
docker build -t 你的仓库/fapiao-baoxiao:latest .
docker push 你的仓库/fapiao-baoxiao:latest
```

---

## 配置项

全部通过环境变量配置，不设就是「本机模式」。

| 变量 | 默认 | 说明 |
|---|---|---|
| `FB_SERVER` | 空 | 填 `1` = 服务器模式：绑 `0.0.0.0`、不自动开浏览器、空闲不退出、**要求登录** |
| `FB_HOST` | `127.0.0.1`（服务器模式为 `0.0.0.0`） | 监听地址 |
| `FB_PORT` | `8766` | 监听端口 |
| `FB_DATA_DIR` | 程序目录 | 数据目录：台账、配置、缓存、输出、回收站、备份都在这 |
| `FB_SOURCE` | 程序目录 | 默认票据文件夹（容器里通常挂到 `/票据`） |
| `FB_PASSWORD` | 空 | 访问口令。**首次启动会用它建管理员账号 `admin`**；也可用于「用户名留空 + 只填口令」的兜底登录。留空 = 不要登录（只在完全可信的内网这么干） |
| `FB_CHROME` | 自动查找 | Chromium / Edge 可执行文件（打印报销单 PDF 用） |
| `TZ` | 系统时区 | 容器时区，填 `Asia/Shanghai` 否则单据时间差 8 小时 |

---

## 数据在哪、怎么备份

数据全在挂载出来的 `data/` 目录里（容器内的 `/data`）：

| 位置 | 内容 |
|---|---|
| `data/发票台账.db` | **台账真身**（SQLite）。账号、权限、审计记录也在里面。**备份就拷它** |
| `data/发票台账.xlsx` | 台账的**导出快照**，双击用 Excel 看 / 打印 |
| `data/输出/` | 生成的报销单、合并 PDF、统计 Excel |
| `data/台账备份/` | 删记录 / 清空 / 导入前的自动备份（`.db` 和 `.xlsx` 各一份） |
| `data/缓存/` | 解析缓存、手机号姓名对照表、出行人核对记录（可删，删了重新解析） |
| `data/回收站/` | 删掉的生成文件挪到这（删错能捞回来） |
| `data/launch.log` | 运行日志，出问题先看它 |
| `票据/上传/` | 从本机传上去的票据副本 |

**备份**：整目录拷 `data/`（或给这个文件夹配同步 / 快照任务）。

> 别在容器跑着的时候**只拷 `发票台账.db` 一个文件** —— SQLite 开了 WAL，部分数据可能还在
> `发票台账.db-wal` 里，单独拷会得到不完整的备份。要么整目录拷，要么先 `docker compose stop`。

---

## 角色与权限

「设置 → 账号与权限」（只有管理员看得到）里给每人一个角色：

| 角色 | 看台账 | 入库 | 出报销单 | 改台账 | 删 / 清空 | 改设置 | 管账号 |
|---|---|---|---|---|---|---|---|
| 管理员 | 是 | 是 | 是 | 是 | 是 | 是 | 是 |
| 财务审核员 | 是 | 是 | 是 | 是 | 是 | 否 | 否 |
| 业务人员 | 是 | 是 | 是 | 是 | 否 | 否 | 否 |
| 只读用户 | 是 | 否 | 否 | 否 | 否 | 否 | 否 |

界面上藏起来的按钮只是「不显眼」，真正卡权限的地方在后端 —— **每个接口都会再查一次**，绕过界面直接发请求也越不了权。

**审计日志**：「设置 → 台账数据库 → 审计日志」。每条含时间、谁、来源 IP、动作、改前、改后；登录失败和越权尝试也在里面。
要停谁的号：那一行点**停用**或**强制退出**，对方已登录的会话立刻失效（不用重启容器）。

---

## 升级到新版本

```bash
docker compose pull
docker compose up -d
```

**数据不受影响** —— 台账、报销单都在 `data/` 卷里，跟镜像无关。

改了本机代码想推上来（GitHub Actions 会自动重新构建镜像，约 3~5 分钟）：

```bash
git add . && git commit -m "改了 xxx" && git push
```

看构建进度：仓库页 → **Actions**。构建出来的镜像每次都会覆盖 `:latest`，同时留一份以提交号命名的版本，改坏了可以退回上一个 tag。

---

## 常见问题

**拉不到镜像 / `unauthorized`**
镜像是私有的，NAS 必须先登录一次。按这个顺序排查：

```bash
whoami                                  # ① 记住你现在是哪个用户（root 还是别人）
docker logout ghcr.io                   # ② 清掉可能残留的错误凭据
echo '你的Token' | docker login ghcr.io -u ZJM117 --password-stdin
                                        # ③ 看到 Login Succeeded 才算过
docker pull ghcr.io/zjm117/fapiao-baoxiao:latest
```

报错文案能直接区分原因：

| 报错里出现 | 含义 | 怎么办 |
|---|---|---|
| `unauthorized` / `authentication required` | **根本没带凭据**（匿名请求） | 十有八九是「登录的用户」和「执行 docker 的用户」不是同一个人：登录的用户凭据存在 `~/.docker/config.json`，用 `sudo` 或 `su root` 后读的是 `/root/.docker/config.json`。**在同一个用户下重新 `docker login`。** |
| `denied` / `permission_denied` / `does not match expected scopes` | 凭据带了，但**Token 权限不够** | 重新建一个 classic Token，勾上 **`read:packages`**，再登录。 |

> 生成 Token 的直链（已预勾 `read:packages`）：<https://github.com/settings/tokens/new?scopes=read:packages&description=nas-pull>
> 必须用 classic Token（`ghp_…`）。设备码 / OAuth 令牌（`gho_…`）即使是本人账号，**拉私有镜像也会被拒**。

另外 **ghcr 路径必须全小写**：`ghcr.io/zjm117/...`（用户名 `ZJM117` 带大写，镜像路径里要小写）。

**页面上显示的时间差 8 小时**
compose 里的 `TZ: "Asia/Shanghai"` 丢了，加回去重建容器。

**报销单 PDF 生成失败 / 中文全是方块**
镜像里缺 Chromium 或中文字体（打印 PDF 靠它）。本仓库的 Dockerfile 已经装好了，自己改过的话确认 `chromium` 和 `fonts-noto-cjk` 还在，然后重新构建。

**忘了口令 / 忘了管理员密码**
改 `docker-compose.yml` 里的 `FB_PASSWORD` 再 `docker compose up -d`（数据不动）。
（那行若写成 `${FB_PASSWORD}` 占位符，就去改**同目录的 `.env`**。）
⚠️ 这只影响「访问口令」这条路；各人账号密码存在数据库里，不受影响。
**管理员密码也忘了**：登录时**用户名留空、只填访问口令**就能进去，再在设置页改回来。

**系统代理 / 内网拉镜像失败**
`docker login` 和 `docker pull` 失败多半是 NAS 的出网问题。给 NAS 的 Docker 配上能用的 DNS 或代理再试。

**端口 8766 被占用**
改 compose 里 `ports` 左边那个数，比如 `- "18766:8766"`。

**同一时间几个人一起改会不会冲突**
不会。台账是 SQLite，写操作都走数据库事务：同时入库、同时改**不同**的行互不影响；同一行被两人同时改，只有先提交的生效，另一个会提示「这行刚被改过，请重新加载」，不会静默丢数据。

**在 Excel 里手改台账，系统会跟着变吗**
不会。Excel 现在是**导出快照**，不是真身。程序重新导出前会先把你的改动另存成
`台账备份/发票台账_手改_<时间>.xlsx`，不会静默覆盖。想把改动并回系统：用「设置 → 台账数据库 → 从 Excel 导入」。

**想把本机的台账搬到 NAS**
把本机 `data/` 下的 `发票台账.db`（有 `-wal` / `-shm` 就一起）拷到 NAS 的 `fapiao/data/`，重启容器。这是台账真身，带过去跟本机一模一样。

---

## 仓库里没有你的数据

**本仓库只有代码。** 真实台账（`发票台账.db`）、票据、缓存（含手机号姓名对照表）、生成的报销单都在
`.gitignore` 里，不会跟着上来；构建镜像时 `.dockerignore` 还会再挡一层。
推送前扫一眼 `git status`，确认列表里没有这些东西。

---

## 开发与测试

```bat
python 测试\diag_deploy_first_run.py   :: 模拟容器首次启动（建库 / 迁移 / 建管理员 / 登录）
python 测试\diag_db.py                 :: 数据层：并发入库、乐观锁、软删除、审计
python 测试\diag_auth.py               :: 登录、角色卡口、越权拦截、会话失效
python 测试\diag_files_guard.py        :: 静态文件服务的暴露面
python 测试\diag_server_mode.py        :: 服务器模式（Docker 那套环境变量）
python 测试\diag_crud.py               :: 台账增删改查
python 测试\diag_merge.py              :: 报销单合并
python 测试\diag_jy.py                 :: 模板二版式
python 测试\diag_upload.py             :: 上传本机文件夹
```

全部在 `%TEMP%` 副本里跑，无窗口，不碰真实数据。

### 代码结构

| 文件 | 职责 |
|---|---|
| `app_web.py` | 界面主程序：HTTP 服务、登录鉴权、接口分发 |
| `app_paths.py` | 路径常量 + `FB_*` 环境变量（本机 / 服务器两种模式的切换点） |
| `db.py` | SQLite 数据层：台账、用户、角色、审计、后台任务 |
| `ledger.py` | 台账读写、导入导出、备份 |
| `invparse.py` | 票据解析（PDF / OFD / XML） |
| `attachment.py` | 发票与行程单配对 |
| `traveler.py` | 出行人识别（三层兜底） |
| `report_pdf.py` | 报销单生成与合并 |
| `gui/` | 前端界面 |

更多细节见 `Docker部署说明.md`（NAS 部署与排障）和 `GitHub与镜像部署说明.md`（推代码 / 自动构建镜像）。
