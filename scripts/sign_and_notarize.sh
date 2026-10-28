#!/usr/bin/env bash
# 对已构建的 .app 做 Developer ID 签名 + Apple 公证 + staple
#
# 前置条件：
#   1. 已加入 Apple Developer Program，并在钥匙串安装 "Developer ID Application" 证书
#   2. 在 https://appleid.apple.com 生成 App 专用密码
#
# 用法：
#   export CERT_ID="Developer ID Application: 你的名字 (TEAMID)"
#   export APPLE_ID="you@example.com"
#   export APPLE_TEAM_ID="TEAMID"
#   export APPLE_APP_PASSWORD="xxxx-xxxx-xxxx-xxxx"
#   ./scripts/sign_and_notarize.sh "dist/Lsy的Excel-Tool.app"
set -euo pipefail

APP="${1:?用法: $0 <path-to-.app>}"
: "${CERT_ID:?请设置 CERT_ID（Developer ID Application 证书名称）}"
: "${APPLE_ID:?请设置 APPLE_ID}"
: "${APPLE_TEAM_ID:?请设置 APPLE_TEAM_ID}"
: "${APPLE_APP_PASSWORD:?请设置 APPLE_APP_PASSWORD（App 专用密码）}"

if [ ! -d "$APP" ]; then
    echo "找不到 $APP" >&2
    exit 1
fi
ZIP="${APP%.app}.zip"

echo "[1/4] 签名..."
codesign --deep --force --options runtime --timestamp --sign "$CERT_ID" "$APP"
codesign --verify --deep --strict --verbose=2 "$APP"

echo "[2/4] 打包并提交公证..."
rm -f "$ZIP"
ditto -c -k --keepParent "$APP" "$ZIP"
xcrun notarytool submit "$ZIP" \
    --apple-id "$APPLE_ID" \
    --team-id "$APPLE_TEAM_ID" \
    --password "$APPLE_APP_PASSWORD" \
    --wait

echo "[3/4] staple 公证票据..."
xcrun stapler staple "$APP"

echo "[4/4] 重新打包（包含 staple 后的 .app）..."
rm -f "$ZIP"
ditto -c -k --keepParent "$APP" "$ZIP"

spctl --assess --type execute --verbose=2 "$APP" || \
    echo "提示：spctl 检查未通过，可把 .app 移到 /Applications 后重试，或检查证书/公证是否成功。"

echo
echo "完成：$APP"
echo "分发文件：$ZIP"
