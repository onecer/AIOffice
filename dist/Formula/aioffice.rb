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
  version "1.23.0"
  license "Apache-2.0"

  # Each platform downloads the matching prebuilt single-file binary from the
  # v#{version} GitHub release and installs it as `aioffice`.
  on_macos do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-arm64"
      # asset: aioffice-mac-arm64 (v1.23.0 SHA256SUMS)
      sha256 "2e32d78b77897bf5a04e089251798156e0a5cf71dfcfc0c2785c4c78dbe0546f"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-mac-x64"
      # asset: aioffice-mac-x64 (v1.23.0 SHA256SUMS)
      sha256 "8a0a59fe654d8920e2dbbfca197f5143d9cb026b9de1cb2c6e990dd32dafaf1f"
    end
  end

  on_linux do
    on_arm do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-arm64"
      # asset: aioffice-linux-arm64 (v1.23.0 SHA256SUMS)
      sha256 "57083c8eacc047122d55491aed75fb29399b326f02eecfdee9ff06cfdc504a87"
    end
    on_intel do
      url "https://github.com/onecer/AIOffice/releases/download/v#{version}/aioffice-linux-x64"
      # asset: aioffice-linux-x64 (v1.23.0 SHA256SUMS)
      sha256 "0938f59f814d036d5146b2e6221cb5530621cf3d0088b2380535b91d39788ce1"
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
