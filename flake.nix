{
  description = "CISiphyus dev shell with NixOS Playwright support";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs = { nixpkgs, ... }:
    let
      systems = [
        "aarch64-linux"
        "x86_64-linux"
      ];
    in
    {
      devShells = nixpkgs.lib.genAttrs systems (system: {
        default = import ./shell.nix {
          pkgs = nixpkgs.legacyPackages.${system};
        };
      });
    };
}
