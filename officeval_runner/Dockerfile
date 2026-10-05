ARG PYTHON_BASE_IMAGE=python:3.12-slim-bookworm
FROM ${PYTHON_BASE_IMAGE}


ARG LIBREOFFICE_VERSION=26.2.6
ARG LIBREOFFICE_ARCHIVE_SHA256=fd0e8f8f2408dd2e5b90286e60f3f97cf566ba441cd48cfc5bcc68067303e0bc

ENV DEBIAN_FRONTEND=noninteractive \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    PYTHONUTF8=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/home/officeval \
    CODEX_HOME=/codex-home \
    XDG_RUNTIME_DIR=/run \
    TMPDIR=/tmp \
    PYTHONPATH=/opt/officeval-src

# Keep LibreOffice off the GPU so it never spends time probing OpenGL, Vulkan
# or OpenCL: the container has no GPU, and CPU rendering stays reproducible.
ENV SAL_DISABLEGL=1 \
    SAL_DISABLESKIA=1 \
    SAL_SKIA=raster \
    SAL_DISABLE_OPENCL=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        bash \
        ca-certificates \
        coreutils \
        curl \
        file \
        findutils \
        fontconfig \
        fonts-crosextra-caladea \
        fonts-crosextra-carlito \
        fonts-dejavu-core \
        fonts-liberation2 \
        fonts-noto-cjk \
        fonts-noto-color-emoji \
        fonts-noto-core \
        git \
        grep \
        gzip \
        jq \
        less \
        libcups2 \
        libdbus-1-3 \
        libglib2.0-0 \
        libx11-xcb1 \
        libxinerama1 \
        libxml2-utils \
        poppler-utils \
        procps \
        qpdf \
        ripgrep \
        sed \
        tar \
        tree \
        unzip \
        zip \
    && rm -rf /var/lib/apt/lists/*

RUN libreoffice_series="${LIBREOFFICE_VERSION%.*}" \
    && libreoffice_url="https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/deb/x86_64/LibreOffice_${LIBREOFFICE_VERSION}_Linux_x86-64_deb.tar.gz" \
    && curl --fail --location --silent --show-error \
        --retry 5 --retry-all-errors \
        --output /tmp/libreoffice.tar.gz \
        "${libreoffice_url}" \
    && echo "${LIBREOFFICE_ARCHIVE_SHA256}  /tmp/libreoffice.tar.gz" | sha256sum -c - \
    && mkdir /tmp/libreoffice \
    && tar -xzf /tmp/libreoffice.tar.gz -C /tmp/libreoffice \
    && apt-get update \
    && apt-get install -y --no-install-recommends /tmp/libreoffice/LibreOffice_*/DEBS/*.deb \
    && ln -s "/opt/libreoffice${libreoffice_series}/program/soffice" /usr/local/bin/libreoffice \
    && ln -s "/opt/libreoffice${libreoffice_series}/program/soffice" /usr/local/bin/soffice \
    && fc-cache -f \
    && soffice --headless --version \
    && mkdir -p /tmp/libreoffice-smoke/profile /tmp/libreoffice-smoke/output \
    && printf 'LibreOffice smoke test\n' > /tmp/libreoffice-smoke/input.txt \
    && soffice --headless --norestore \
        "-env:UserInstallation=file:///tmp/libreoffice-smoke/profile" \
        --convert-to pdf \
        --outdir /tmp/libreoffice-smoke/output \
        /tmp/libreoffice-smoke/input.txt \
    && pdfinfo /tmp/libreoffice-smoke/output/input.pdf >/dev/null \
    && rm -rf \
        /tmp/libreoffice \
        /tmp/libreoffice-smoke \
        /tmp/libreoffice.tar.gz \
        /var/lib/apt/lists/*

WORKDIR /opt/officeval-src
COPY requirements-container.txt ./
RUN --mount=type=cache,id=officeval-pip,target=/var/cache/pip,sharing=locked \
    PIP_CACHE_DIR=/var/cache/pip python -m pip install \
        --timeout 600 \
        --retries 10 \
        -r requirements-container.txt \
    && PIP_CACHE_DIR=/var/cache/pip python -m pip install \
        --timeout 600 \
        --retries 10 \
        --no-deps \
        rapidocr-onnxruntime==1.4.4 \
    && python -c "import cv2, fitz, numpy, pdfplumber, rapidocr_onnxruntime, reportlab, scipy" \
    && mkdir -p /workspace /home/officeval /codex-home \
    && chmod 0700 /home/officeval /codex-home
# A pip cache would fill the Agent's 64 MiB HOME before the install finishes.
ENV PIP_NO_CACHE_DIR=1
COPY src ./src

WORKDIR /workspace
ENTRYPOINT ["python", "-P", "-m", "src.runtime.worker"]
