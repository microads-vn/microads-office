# Code signing policy

Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).

## What is signed

Only binaries built from this repository by the GitHub Actions workflow [`release.yml`](.github/workflows/release.yml) are submitted for signing:

* `MicroadsOffice.exe` — the Microads Office suite (Word, Excel, PDF, hub, installer), published as `MicroadsOffice_Setup_<version>.exe`
* `MicroadsOffice_Setup.exe` — the web installer

No third-party or upstream binaries are signed. The only third-party binary shipped inside the suite is PDFium (see
[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)), embedded unmodified.

## Team roles

| Role | Members |
|---|---|
| Committers and reviewers | Members of the [microads-vn](https://github.com/microads-vn) GitHub account, Microads Technology Company Limited |
| Approvers | Owner of the [microads-vn](https://github.com/microads-vn) GitHub account (Microads Technology Company Limited) |

Every signing request is approved manually by an approver. All team members use multi-factor authentication for GitHub and SignPath.

## Privacy policy

This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing
or operating it, **except** for the automatic update check described in the privacy policy, which is shown during installation and can be
switched off there (or later in *Help › Settings*): <https://microads.vn/office/privacy>.

The optional **AI assistant** is not connected until the user chooses a service (their ChatGPT account, their own Claude or OpenAI
API key, their own server, or Claude Desktop). It sends a request only when the user clicks *Send* or a quick action, directly to the
service the user chose — never through Microads servers — and the user can untick "include the open document" at any time.

## Release integrity

In addition to the Authenticode signature, every release is signed with the Microads ECDSA P-256 release key; the apps, the web installer and
the update server verify that signature before installing anything.

Publisher: Microads Technology Company Limited (Công ty TNHH Công nghệ Microads), Ho Chi Minh City, Vietnam — <https://microads.vn>.
