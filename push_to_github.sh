#!/bin/bash
# GitHub push 스크립트
# 사용법: ./push_to_github.sh YOUR_GITHUB_TOKEN

TOKEN=$1
if [ -z "$TOKEN" ]; then
  echo "사용법: ./push_to_github.sh <GitHub_PAT_Token>"
  echo "GitHub Settings → Developer settings → Personal access tokens → Generate new token"
  echo "필요 권한: repo (Full control)"
  exit 1
fi

git remote set-url origin "https://${TOKEN}@github.com/hjyun9731-gif/misugeum.git"
git push -u origin main
echo "✅ Push 완료!"
