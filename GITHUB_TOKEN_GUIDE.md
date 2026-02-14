# GitHub Personal Access Token 创建指南

由于GitHub已停止支持密码认证，你需要创建Personal Access Token来访问GitHub API。

## 创建步骤：

1. 登录你的GitHub账户 (https://github.com)

2. 点击右上角的头像，选择 "Settings"

3. 在左侧菜单中，向下滚动并点击 "Developer settings"

4. 点击 "Personal access tokens" → "Tokens (classic)"

5. 点击 "Generate new token" → "Generate new token (classic)"

6. 填写以下信息：
   - **Note**: 输入一个描述性名称，如 "WpsAgent Deploy"
   - **Expiration**: 选择过期时间（建议选择 90 days 或 No expiration）
   - **Select scopes**: 勾选以下权限：
     - `repo` (完整的仓库访问权限)
     - `workflow` (如果需要GitHub Actions)

7. 点击 "Generate token" 按钮

8. **重要**: 复制生成的token（格式类似：ghp_xxxxxxxxxxxx）
   - 这个token只会显示一次，请妥善保存

## 使用Token：

创建好token后，请告诉我你的token，我将使用它来创建GitHub仓库并上传代码。

**注意**: Token具有完整的仓库访问权限，请妥善保管，不要分享给他人。