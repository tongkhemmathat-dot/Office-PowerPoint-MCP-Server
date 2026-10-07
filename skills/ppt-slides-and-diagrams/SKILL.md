---
name: ppt-slides-and-diagrams
description: Build PowerPoint decks and system diagrams (cloud, on-prem hardware, software, network) with the Office-PowerPoint-MCP-Server tools. Use when a system engineer asks for slides, a design/proposal deck, or an architecture, deployment or network diagram as .pptx.
---

# PowerPoint slides and system diagrams (ppt MCP)

Needs the `ppt` MCP server v2.1.0+. All positions are inches; a 16:9 slide is 13.333 x 7.5.

## Pick the path
- **Text/table/chart slides** -> "Deck workflow".
- **Architecture / network / deployment diagram** -> "Diagram workflow". A deck usually mixes both.
- Ask only what changes the result: audience, purpose, slide count, language, template file. Otherwise assume 16:9.

## Deck workflow
1. `create_presentation(id, slide_size="16:9")`, or `create_presentation_from_template` when the team has a template (then use its layouts).
2. Outline first: title, context/problem, solution, architecture, details (tables), risks, next steps. One message per slide; the title states the takeaway ("Peak load needs 4 app nodes"), not the topic.
3. `add_slide(layout_index=...)`. Default-template layouts: 0 title, 1 title+content, 2 section, 3 two content, 4 comparison, 5 title only, 6 blank. Check with `get_slide_info`.
4. Fill: `populate_placeholder`, `add_bullet_points` (<= 6 bullets, short), `add_table` for specs/sizing/BoM/port lists (`data` is a list of rows of strings and must have exactly `rows` rows; `header_row`), `add_chart` for numbers, `manage_text` for free text boxes, `manage_image` for pictures.
5. `list_slide_templates` / `create_slide_from_template` and `apply_professional_design` give styled layouts; `auto_generate_presentation` is only a rough draft generator. Read each tool's schema for argument formats before using it.
6. `set_core_properties` (title/author), then `save_presentation`.
7. Render every slide (below) and fix overflow before reporting done.

## Diagram workflow
1. Decide the view: **context** (system + externals), **logical** (components and flows), **physical/deployment** (hosts, networks, zones). Use one slide per view; split a slide that needs more than ~20 elements.
2. Sketch a grid: flow left->right or top->bottom, one row/column per tier, 0.5-1 in gaps.
3. Send everything in **one** `add_diagram_elements` call. Order = z-order: group boxes first, then icons/shapes and text, then connectors last.
4. `render_slide(return_image=true)`, look, fix (`manage_shapes` or resend), repeat.
5. Add a small legend when colours or line styles carry meaning.

### Elements
```json
{"type":"shape","shape_type":"rectangle","left":2,"top":2,"width":7,"height":4,"fill_color":[247,251,243],"line_color":[122,161,22],"line_width":2}
{"type":"shape","shape_type":"cylinder","left":5,"top":3,"width":0.8,"height":1,"fill_color":[230,242,252],"line_color":[0,115,187],"text":"DB","font_size":10}
{"type":"icon","name":"EC2","left":4,"top":3,"size":0.6,"label":"App server\n(2 nodes)","label_width":1.5}
{"type":"text","text":"DMZ","left":2.1,"top":2.05,"width":2,"height":0.25,"bold":true,"font_size":10}
{"type":"connector","start_x":3,"start_y":3.3,"end_x":4,"end_y":3.3,"arrow_end":"arrow","color":[237,112,20],"line_width":2}
```
Bad elements are skipped and listed under `errors`; the rest are still added. Resend only the failed ones.

### Shapes when there is no icon (hardware, software, network)
`rectangle`/`rounded_rectangle` server, service or component; `cylinder` database/storage; `cube` appliance/node; `cloud` internet/WAN; `hexagon`/`octagon` firewall or gateway; `chevron` and the `*_arrow` shapes for process flow; `flowchart_*` for decisions. Put the name inside via `text` and keep the model/spec on a second line ("DB-01\nPostgreSQL 16, 8 vCPU").

### Icons
- AWS: `python scripts/fetch_aws_icons.py` fills `icons/aws/` (AWS icon terms apply; not bundled). Common names: Users, Internet, Client, Route53, CloudFront, WAF, Shield, ElasticLoadBalancingApplicationLoadBalancer, EC2, EC2Instance, EC2AutoScaling, Lambda, RDS, DynamoDB, EFS, FSx, SimpleStorageService, Backup, CloudWatch, CertificateManager, SecretsManager, IAMIdentityCenter, VPCEndpoints. Search others with `list_icons(query=...)`.
- Docker / Kubernetes: `python scripts/fetch_container_icons.py` fills `icons/k8s/` (official K8s set, Apache-2.0: k8s-pod, k8s-deployment, k8s-service, k8s-ingress, k8s-configmap, k8s-secret, k8s-persistent-volume-claim, k8s-namespace, k8s-node, k8s-control-plane, k8s-etcd, ... use `list_icons(query="k8s")`) and `icons/docker/` (logo-docker, logo-docker-wordmark, logo-kubernetes, logo-helm, logo-podman). Docker has no official object icons: use se-container, se-container-image, se-registry, se-database (volume). Follow the Docker/Kubernetes trademark guidelines for logos.
- Built-in generic set `se-*` (assets/icons/se): se-container, se-container-image, se-registry, se-server, se-server-rack, se-app-server, se-web-server, se-gis-server, se-db-server, se-database, se-nas, se-vm, se-esx-host, se-network-switch, se-firewall, se-internet, se-user, se-user-web, se-user-mobile, se-workstation.
- Visio stencils/diagrams (the team's own icons): `python scripts/extract_visio_icons.py <file-or-folder>` writes `icons/visio/visio-<master name>.png` (pictures read from the file; vector masters exported through Visio, needs Windows + Visio + pywin32). Keep vendor icons local, never commit them. Old binary `.vss`/`.vsd` must be re-saved as `.vssx`/`.vsdx` in Visio first.
- Azure, GCP and other vendors: no set ships with this server. Point `PPT_ICON_PATH` at a folder of PNG/JPG icons the team is licensed to use (SVG needs `cairosvg`), or use shapes. Never claim an icon exists without `list_icons` confirming it.
- Icons are 64 px PNGs: keep 0.6-0.8 in.

### Conventions
- Colour by meaning, consistently: e.g. edge/public green, internal/private blue, data orange or purple, external grey. Don't use colour alone; label it.
- Arrows: solid = synchronous/primary flow, `dash_style="dash"` = async/backup/optional, orange 2 pt for user traffic, blue 1.5 pt internal. Label protocol/port on key links ("HTTPS 443"), number the steps for sequences.
- Container titles go in a corner clear of icons. Start arrows 0.05-0.1 in outside icon edges and keep them off labels.
- Labels: 9-10 pt, <= 2 short lines, name first, role second. Minimum 9 pt anywhere on a slide.
- Redundancy: draw AZ/site/node pairs explicitly (A/B) instead of writing "HA" in a label.
- Thai text: set `font_name` (e.g. "Tahoma") on text elements and check it in the render; label boxes are narrow, so widen `label_width`.

## Verify
`render_slide` needs PowerPoint (Windows) or LibreOffice; without either, say the layout is unchecked. Look for overlaps, clipped text, arrows crossing labels, unreadable sizes.

## Editing
`manage_shapes(operation="list")` -> index, name, position. `move` / `resize` / `delete` / `bring_to_front` / `send_to_back` by `shape_index`; a delete shifts later indexes, so re-list.

## Gotchas
- `create_presentation` with an existing `id` replaces that presentation.
- 4:3 is the default size; pass `slide_size="16:9"`.
- Restart the MCP server after upgrading it, or old tool behaviour persists.
- Don't put real hostnames, IPs or credentials on slides unless the user supplied them for that purpose.
