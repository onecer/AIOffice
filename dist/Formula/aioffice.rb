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
  version "1.11.0"
  license "Apache-2.0"

  # Each platform downloads the matching prebuilt single-file binary from the
  # v#{version} GitHub release and installs it as `aioffice`.
  on_macos do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-arm64"
      # asset: aioffice-mac-arm64
      # TODO(human): fill from v1.11.0 SHA256SUMS. Example value below is from v1.5.0.
      sha256 "63a5d987013ed77561ee5e50d4a8e4f66ae81ec668cb0a5e99cc2b5e0c6aee6c"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-x64"
      # asset: aioffice-mac-x64
      # TODO(human): fill from v1.11.0 SHA256SUMS. Example value below is from v1.5.0.
      sha256 "b8bc8d2467704ae0bab6fa6faffc4e4d92e80998c0d554d90fd48ab2eca63fe1"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-arm64"
      # asset: aioffice-linux-arm64
      # TODO(human): fill from v1.11.0 SHA256SUMS. Example value below is from v1.5.0.
      sha256 "c73b06d3e1ad97989088bb4ab264f83bab8804e0b8e14d1c27174e533d00b5c0"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-x64"
      # asset: aioffice-linux-x64
      # TODO(human): fill from v1.11.0 SHA256SUMS. Example value below is from v1.5.0.
      sha256 "842a55b1fb4f753a61ae5e534afb1a364cf331b49b83a9a305c79270bfebd570"
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
