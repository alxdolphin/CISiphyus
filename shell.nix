{ pkgs ? import <nixpkgs> {} }:

let
  # WHY: pip playwright revision must match nixpkgs playwright-driver.browsers
  playwrightPipVersion = pkgs.python311Packages.playwright.version;
in

pkgs.mkShell {
  packages = with pkgs; [
    python311
    stdenv.cc.cc.lib
    playwright-driver.browsers
  ];

  shellHook = ''
    export PLAYWRIGHT_BROWSERS_PATH=${pkgs.playwright-driver.browsers}
    export PLAYWRIGHT_SKIP_VALIDATE_HOST_REQUIREMENTS=true
    export LD_LIBRARY_PATH=${pkgs.stdenv.cc.cc.lib}/lib''${LD_LIBRARY_PATH:+:}$LD_LIBRARY_PATH

    if [ ! -d .venv ]; then
      python -m venv .venv
    fi
    .venv/bin/python -m pip install -q -e '.[dev]' 'playwright==${playwrightPipVersion}'
    source .venv/bin/activate
  '';
}
