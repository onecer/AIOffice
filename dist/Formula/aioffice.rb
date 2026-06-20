# Homebrew formula for AIOffice — the AI-native Office CLI + MCP.
#
# Tap usage (after the human creates onecer/homebrew-tap and adds this file as
# Formula/aioffice.rb in that repo):
#
#   brew install onecer/tap/aioffice
#
# or, in two steps:
#
#   brew tap onecer/tap
#   brew install aioffice
#
# -------------------------------------------------------------------------
# RELEASING A NEW VERSION
# -------------------------------------------------------------------------
# 1. Bump `version` below.
# 2. Replace every sha256 with the value from that release's SHA256SUMS:
#       gh release download v<version> -p SHA256SUMS -O - | sort
#    Each line is "<sha256>  <asset-name>"; map asset -> the matching sha256
#    field in this file (see the comment on each line).
# 3. Commit the updated formula to the homebrew-tap repo.
# -------------------------------------------------------------------------
class Aioffice < Formula
  desc "AI-native CLI and MCP for .docx/.xlsx/.pptx, no Office install needed"
  homepage "https://github.com/onecer/AIOffice"
  version "1.14.0"
  license "Apache-2.0"

  # Each platform downloads the matching prebuilt single-file binary from the
  # v#{version} GitHub release and installs it as `aioffice`.
  on_macos do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-arm64"
      # asset: aioffice-mac-arm64 (v1.14.0 SHA256SUMS)
      sha256 "dccc39a26ad50221880d96d2e7de7023f75a83b9c85d708ba775f824005a67e8"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-x64"
      # asset: aioffice-mac-x64 (v1.14.0 SHA256SUMS)
      sha256 "eb5f55eeaa979c988d3f6687cb42be95367a961abb13b27957b470fbd3fb03b5"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-arm64"
      # asset: aioffice-linux-arm64 (v1.14.0 SHA256SUMS)
      sha256 "47cda325c055b4a9793e48d7fc6486ca9aa5f5e2ca0d7db56fb0eae1745961c0"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-x64"
      # asset: aioffice-linux-x64 (v1.14.0 SHA256SUMS)
      sha256 "48c35211b44df13f8b838e2f4e84204cebc27fda31270ce9c0e7ec42f8b0576b"
    end
  end

  def install
    # The release asset is named per-platform; install it under the canonical name.
    asset = if OS.mac?
      Hardware::CPU.arm? ? "aioffice-mac-arm64" : "aioffice-mac-x64"
    else
      Hardware::CPU.arm? ? "aioffice-linux-arm64" : "aioffice-linux-x64"
    end
    bin.install asset => "aioffice"
  end

  test do
    output = shell_output("#{bin}/aioffice version")
    assert_match "\"version\":\"#{version}\"", output
  end
end
