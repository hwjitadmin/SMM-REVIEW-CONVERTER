# smm-review PDF 変換 API（Cloud Run）
# Debian の LibreOffice（Calc のみ・GUIなし）と、その UNO を使える Debian 標準の Python 3 で動かす。
FROM debian:bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive \
    LANG=C.UTF-8 \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    LO_PROFILE_TEMPLATE=/opt/lo-profile \
    HOME=/tmp

# 日本語フォント：IPA フォントは MS ゴシック／明朝と文字幅が同じになるよう作られているため、
# Excel の MS 系フォントの代わりに使う（fonts.conf で対応付け）。メイリオ・游ゴシックは Noto Sans CJK JP。
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-calc-nogui python3-uno \
      python3-flask gunicorn \
      fonts-ipafont-gothic fonts-ipafont-mincho fonts-ipaexfont fonts-noto-cjk \
      fontconfig \
 && rm -rf /var/lib/apt/lists/*

COPY fonts.conf /etc/fonts/local.conf
RUN fc-cache -f

# LibreOffice の初回起動時の初期化をビルド時に済ませ、プロファイルを雛形として保存しておく。
# 地域設定を日本語（ja-JP）にして、Excel の「システムの日付形式」のセルを 2026/09/29 の形で表示させる
# （既定の英語（米国）のままだと 9/29/2026 になる）。
RUN soffice --headless --terminate_after_init "-env:UserInstallation=file://${LO_PROFILE_TEMPLATE}" \
 && sed -i 's#</oor:items>#<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>ja-JP</value></prop></item></oor:items>#' \
      "${LO_PROFILE_TEMPLATE}/user/registrymodifications.xcu" \
 && grep -q 'ooSetupSystemLocale' "${LO_PROFILE_TEMPLATE}/user/registrymodifications.xcu" \
 && chmod -R a+rX "${LO_PROFILE_TEMPLATE}"

WORKDIR /app
COPY app/ /app/

RUN useradd --system --no-create-home converter
USER converter

# 1インスタンスで同時に1件だけ変換する（LibreOffice はメモリを多く使うため。Cloud Run 側も concurrency=1）
CMD exec gunicorn --bind ":${PORT:-8080}" --workers 1 --threads 1 --timeout 120 main:app
