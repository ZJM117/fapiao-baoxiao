# 放到 GitHub 上 + 用镜像部署到 NAS

一句话：**代码推到 GitHub，GitHub 免费帮你构建镜像，NAS 只负责拉下来跑**。
以后你改了代码，`git push` 一下，NAS 上 `docker compose pull && docker compose up -d` 就更新完了 ——
不用在 NAS 上装编译环境，也不用等它慢慢构建。

```
你的电脑                 GitHub                        你的 NAS
────────              ─────────────                 ─────────────
git push   ───────►   仓库（私有/公开）
                       └─ Actions 自动构建  ───────►  ghcr.io 上的镜像
                                                       │
                                        docker compose pull ─┘
                                                       ▼
                                                容器跑起来（8766）
```

---

## 一、为什么要绕这一圈（直接 build 不行吗）

行，但两个做法差别不小：

| | GitHub 帮你构建（本文这法） | NAS 上现构建 |
|---|---|---|
| NAS 需要什么 | 只要能拉镜像 | 要能连 Debian 源、装 chromium + 中文字体 |
| 第一次 | 几十秒拉完 | 3~5 分钟，还可能卡在 `apt-get` |
| 以后更新 | `pull` + `up -d` | 每次都要重新构建 |
| 万一构建失败 | 在 GitHub 上看得见完整报错 | 只能在 NAS 上看日志猜 |

---

## 二、开始前准备

1. 一个 **GitHub 账号**（没有就注册一个，免费）。
2. 电脑上装好 **Git**（你这台已经装了：`git version 2.55.0`）。
3. 想好仓库名：**建议就叫 `fapiao-baoxiao`** ——
   下面所有命令、配置文件都按这个名字写的，叫别的就得跟着改镜像名。

> 仓库**建议设为 Private（私有）**：这是内部工具，代码里会提到公司名、字段口径这些东西。
> 私有仓库照样能用 GitHub Actions 构建镜像，免费额度（每月 2000 分钟）对这点构建量绰绰有余。

---

## 三、把代码推上去

### 1. 先在 GitHub 网页上建一个空仓库

打开 <https://github.com/new>：

- Repository name 填 `fapiao-baoxiao`
- 选 **Private**
- ⚠️ **不要**勾 "Add a README file"、不要选 .gitignore 模板（我们要用自己的）
- 点 Create repository

### 2. 在本机项目目录里执行

打开 **cmd**（不是这里），逐条粘。

> 📌 本地仓库**已经帮你准备好了** —— `git init`、`git add`、分支名 `main` 都做过了，
> 也确认过 **53 个待提交文件里没有任何数据文件**（台账、备份、缓存、手机号对照表都没有）。
> 所以从 `commit` 开始就行。

```bat
cd /d E:\桌\个人文件同步\workbuddy\发票报销

:: 第一次用 git 的话，先填一下提交身份（只需做一次）
git config --global user.name "你的名字"
git config --global user.email "你的邮箱"

:: 最后扫一眼要提交什么（不该出现的见下面「提交前必看」）
git status

git commit -m "发票报销工具：多人内部使用版"

git remote add origin https://github.com/ZJM117/fapiao-baoxiao.git
git push -u origin main
```

（换个电脑要重做时：先 `git init`、`git branch -M main`、`git add .`，再往下走。）

- 第一次 `commit` 如果提示要填身份，就执行一次：
  ```bat
  git config --global user.name "你的名字"
  git config --global user.email "你的邮箱"
  ```
- `push` 时会弹出一个窗口让你登录 GitHub（Git for Windows 自带的凭据管理器），
  选 **Sign in with your browser** 最简单。**不要**去输账号密码 —— GitHub 早就不收密码了。

### 提交前必看（`git status` 里不该出现的东西）

`.gitignore` 已经把下面这些排掉了，但**自己扫一眼最保险**，出现任何一个都先停下：

- `发票台账.db` / `发票台账.db-wal` —— **台账真身**：全部发票数据 + 账号 + 审计
- `发票台账.xlsx` / `台账备份/` —— 台账快照与备份
- `缓存/`（里面有手机号姓名对照表）、`输出/`（生成过的报销单）
- `config.json`、`*.log`、`票据/`
- `*预览*.txt` —— 里面是真实人名

> 万一已经推上去了：先在 GitHub 上把仓库**删掉重建**（改 `.gitignore` 再重推），
> 别指望「提交一个删除」能把历史里的数据抹掉。

---

## 四、GitHub 自动帮你构建镜像

推上去之后**不用做任何事**，Actions 自己就跑起来了。

**看构建进度**：仓库页 → 顶部 **Actions** 标签 → 点那条 "构建镜像并推送到 ghcr.io"。
绿色 ✓ 就是好了，红 × 就点进去看哪一步报错（通常是 Dockerfile 里那一步）。

**镜像在哪**：仓库页右侧 **Packages**（或 <https://github.com/ZJM117?tab=packages>），
名字是 `fapiao-baoxiao`。

**构建好的镜像地址**：

```
ghcr.io/zjm117/fapiao-baoxiao:latest
```

每次 push 都会覆盖 `:latest`，同时留一份以提交号命名的（比如 `:a1b2c3d…`）——
以后哪次改坏了，把 NAS 上的 tag 换成上一个提交号就能回滚。

---

## 五、NAS 上拉镜像并启动

### 1. 镜像的可见性，先决定一下

GitHub 上这个镜像（Package）默认跟着仓库走 —— **仓库私有，镜像就是私有**。

| 选哪个 | NAS 上要做什么 | 代价 |
|---|---|---|
| **私有**（推荐） | 先 `docker login ghcr.io` 登录一次 | 每个想拉它的设备都要登录 |
| 公开 | 什么都不用做 | ⚠️ 镜像里**含源码**（Dockerfile 里 `COPY . /app`），等于把代码公开了 |

想改成公开：Packages → `fapiao-baoxiao` → Package settings → Change visibility → Public。

### 2. 私有镜像：在 NAS 上登录一次

先做一个 **PAT**（Personal Access Token）：GitHub 右上角头像 → Settings →
Developer settings → Personal access tokens → **Tokens (classic)** → Generate new token (classic)：
- 勾上 **`read:packages`**
- 有效期按自己习惯设，生成后**马上复制**（只显示一次）

然后在 NAS 的 SSH 里（或用绿联的「终端」）：

```bash
docker login ghcr.io
# Username: 你的 GitHub 用户名
# Password: 刚才那个 PAT（不是 GitHub 密码）
```

登录信息会存在 NAS 上，除非重新登录否则一直有效。

### 3. 改 `docker-compose.yml` 一处

把镜像地址里的用户名换成你自己的：

```yaml
    image: ghcr.io/zjm117/fapiao-baoxiao:latest
```

（其余照旧：`FB_PASSWORD` 改口令、`./票据:/票据` 指票据文件夹、端口 8766。）

### 4. 起服务

```bash
cd /vol1/docker/fapiao      # 你放项目的目录（只需要 docker-compose.yml 这一个文件！）
docker compose pull
docker compose up -d
docker compose logs -f      # 看到「[就绪] 请用浏览器访问」就好了
```

> **用镜像的好处之一**：NAS 上其实**只需要 `docker-compose.yml` 一个文件** ——
> 代码都在镜像里了，不用把整个项目文件夹拷上去。
> 想省事就把项目文件夹里这个文件单独拷到 NAS 上即可。

浏览器打开 `http://NAS的IP:8766`，第一次登录：用户名 `admin` + 你设的 `FB_PASSWORD`。

---

## 六、以后怎么更新

```bat
:: 本机：改完代码
cd /d E:\桌\个人文件同步\workbuddy\发票报销
git add .
git commit -m "改了 xxx"
git push
```

等 GitHub 上 Actions 变成绿色 ✓，然后在 NAS 上：

```bash
cd /vol1/docker/fapiao
docker compose pull
docker compose up -d
```

**数据不会动** —— 台账、报销单都在 `/data` 卷里，跟镜像无关。

---

## 七、常见问题

**Actions 里那一步红了怎么办**
点进去看日志。最常见的是 Dockerfile 里 `apt-get` 那一步拉不到包（Debian 源网络问题）——
重新跑一次通常就好（Actions 页面右上角有 "Re-run jobs"）。

**`docker compose pull` 报 denied / unauthorized**
镜像还是私有的，而 NAS 没登录（或 PAT 过期了）。回到第五节的第 2 步重新 `docker login ghcr.io`。
注意 PAT 要有 **read:packages** 权限。

**构建出来的镜像能给别人用吗**
能，但要知道：**镜像里含源码**。想让别人直接用，就把 Package 设成 Public
（代码仓库仍然可以是私有的，只不过等于把代码一起公开了）。

**GitHub 构建要不要花钱**
私有仓库每月有 2000 分钟免费额度，这个镜像一次构建约 3~5 分钟，怎么用都用不完。

**想改回「NAS 上自己构建」**
把 `docker-compose.yml` 里的 `image:` 那行注释掉、放开 `build: .`，
然后 `docker compose up -d --build`（也可以把整个项目文件夹拷到 NAS 上再这么干）。

**NAS 是 arm 架构（不是 x86）怎么办**
改 `.github/workflows/build-image.yml` 里的 `platforms: linux/amd64,linux/arm64`，重推一次。
构建时间会翻倍。（你的绿联 DXP4800 Plus 是 x86_64，不用管这条。）

**担心代码里混进真实数据**
`.dockerignore` 和 `.gitignore` 各自挡了一层。每次 `git add .` 之后跑一下 `git status` 肉眼确认，
列表见第三节「提交前必看」。
