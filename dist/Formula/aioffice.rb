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
  version "1.27.0"
  license "Apache-2.0"

  # Each platform downloads the matching prebuilt single-file binary from the
  # v#{version} GitHub release and installs it as `aioffice`.
  on_macos do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-arm64"
      # asset: aioffice-mac-arm64 (v1.27.0 SHA256SUMS)
      sha256 "b2a1238d32aa220fe25990e60773ced034caaa3dde25f0052c250fb03f77c10c"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-x64"
      # asset: aioffice-mac-x64 (v1.27.0 SHA256SUMS)
      sha256 "26bc5eb063c1cebcab25494ab463087c5b8c0ce24113934210afc69c8c1818cc"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-arm64"
      # asset: aioffice-linux-arm64 (v1.27.0 SHA256SUMS)
      sha256 "84fd0e871a2187cbefbf5f126987bc86705805d5416b313bcb93e84e7335aeaa"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-x64"
      # asset: aioffice-linux-x64 (v1.27.0 SHA256SUMS)
      sha256 "d684f2b9bc55a1d5fcc86d3aa9333e7f940155d9aa1bbc278fc403a257def861"
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
