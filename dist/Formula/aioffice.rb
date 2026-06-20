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
  version "1.12.0"
  license "Apache-2.0"

  # Each platform downloads the matching prebuilt single-file binary from the
  # v#{version} GitHub release and installs it as `aioffice`.
  on_macos do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-arm64"
      # asset: aioffice-mac-arm64 (v1.12.0 SHA256SUMS)
      sha256 "0d30c51c114e46ecb6d78d2800728d5cd4842b81d15abaafc07b0ef1b679d122"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-x64"
      # asset: aioffice-mac-x64 (v1.12.0 SHA256SUMS)
      sha256 "b28583d01c729ca9723f4ff95848a2b75fda53190dc4782ddee72ca48de51c9f"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-arm64"
      # asset: aioffice-linux-arm64 (v1.12.0 SHA256SUMS)
      sha256 "3606bf0eb4630a069eb2c1e65a96186a8e0542b784dfe2e643621f431dbb6326"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-x64"
      # asset: aioffice-linux-x64 (v1.12.0 SHA256SUMS)
      sha256 "8f112d34a7551f14d8cfbb0e956c44c352d0b0502355040eae1ea53779b38f48"
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
