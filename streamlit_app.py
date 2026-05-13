"""
Streamlit UI for PPTX Slide Audit Agent (JedAI-Powered)
Run:  streamlit run app.py
"""

import os
import sys
import io
import time
import math
import tempfile
import streamlit as st

# ── Import everything from test.py ──
import test as audit

st.set_page_config(page_title="PPTX Slide Audit Agent", page_icon="📊", layout="wide")

st.title("📊 PPTX Slide Audit Agent")
st.caption("Compare training slides against latest product manuals using JedAI")

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR — Configuration
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ Configuration")

    st.subheader("🔐 Credentials")
    username = st.text_input("LDAP Username", value="", key="username")
    password = st.text_input("LDAP Password", type="password", key="password")

    st.subheader("🤖 Model")
    model_mode = st.radio(
        "Model selection",
        ["auto", "Specify model name"],
        index=0,
        help="'auto' picks the first available model from JedAI"
    )
    if model_mode == "Specify model name":
        model_name = st.text_input(
            "Model name",
            value="GCP_claude-opus-4-6",
            help="e.g. GCP_gemini-2.5-pro, AzureOpenAI_gpt-4o, on_prem_openai/gpt-oss-120b"
        )
    else:
        model_name = "auto"

    st.subheader("🔧 Tuning")
    product_mode = st.selectbox(
        "Product / Tool",
        ["generic", "innovus", "genus", "conformal", "tempus", "joules", "pegasus", "modus"],
        index=0,
        help="Select 'conformal' for Conformal/LEC slides (multi-word commands like 'read design'). "
             "'generic' uses underscore-based detection (e.g. set_db, report_timing)."
    )
    enable_phase2 = st.checkbox("Enable Phase 2 (new commands scan)", value=False,
                                 help="Slower — scans manual for commands not in slides")
    chunk_size = st.slider("Manual chunk size", 2000, 12000, 6000, step=500)
    max_tokens = st.slider("Max tokens per response", 2048, 16384, 8192, step=1024)
    max_workers = st.slider("Parallel workers", 1, 10, 5, step=1,
                            help="Number of concurrent API calls. Higher = faster but may hit rate limits")

# ═══════════════════════════════════════════════════════════════════════════
# MAIN — File uploads
# ═══════════════════════════════════════════════════════════════════════════
col1, col2 = st.columns(2)

with col1:
    pptx_file = st.file_uploader("📎 Upload PPTX file", type=["pptx"], key="pptx")

with col2:
    manual_files = st.file_uploader(
        "📎 Upload Manual(s)", type=["pdf", "docx", "txt", "pptx", "md"],
        accept_multiple_files=True, key="manuals"
    )

# ═══════════════════════════════════════════════════════════════════════════
# RUN AUDIT
# ═══════════════════════════════════════════════════════════════════════════
can_run = pptx_file is not None and len(manual_files) > 0 and username and password

if st.button("🚀 Run Audit", disabled=not can_run, type="primary", use_container_width=True):
    if not can_run:
        st.error("Please fill in all required fields: credentials, PPTX file, and at least one manual.")
        st.stop()

    # Apply settings to the audit module
    audit.JEDAI_USERNAME = username
    audit.JEDAI_PASSWORD = password
    audit.JEDAI_MODEL = model_name
    audit.PRODUCT_MODE = product_mode
    audit.ENABLE_PHASE2 = enable_phase2
    audit.CHUNK_SIZE = chunk_size
    audit.MAX_TOKENS = max_tokens
    audit.MAX_WORKERS = max_workers
    audit.AUTH_TOKEN = None  # Force fresh login

    # Save uploaded files to temp directory
    tmp_dir = tempfile.mkdtemp()
    pptx_path = os.path.join(tmp_dir, pptx_file.name)
    with open(pptx_path, "wb") as f:
        f.write(pptx_file.getbuffer())

    manual_paths = []
    for mf in manual_files:
        mp = os.path.join(tmp_dir, mf.name)
        with open(mp, "wb") as f:
            f.write(mf.getbuffer())
        manual_paths.append(mp)

    output_path = os.path.join(tmp_dir, "slide_audit_report.xlsx")

    # ── Estimate runtime ──
    pptx_size_chars = len(pptx_file.getbuffer())
    manual_size_chars = sum(len(mf.getbuffer()) for mf in manual_files)
    est_slide_chunks = max(1, pptx_size_chars // (audit.SLIDE_CHUNK_SIZE * 3))
    est_commands = max(5, pptx_size_chars // 1500)
    est_manual_chunks = max(1, manual_size_chars // (chunk_size * 3))
    est_phase2_calls = est_manual_chunks if enable_phase2 else 0
    total_api_calls = est_slide_chunks + est_commands + est_phase2_calls
    avg_seconds_per_call = 15
    est_seconds = math.ceil(total_api_calls / max_workers) * avg_seconds_per_call + 20
    est_h = int(est_seconds // 3600)
    est_m = int((est_seconds % 3600) // 60)
    est_s = int(est_seconds % 60)
    est_time_str = f"{est_h:02d}:{est_m:02d}:{est_s:02d}"

    st.info(
        f"⏱️ **Estimated runtime:** ~{est_time_str} "
        f"({total_api_calls} API calls ÷ {max_workers} workers × ~{avg_seconds_per_call}s/call)"
    )

    start_time = time.time()

    # Capture console output for the log
    log_area = st.empty()
    progress_bar = st.progress(0, text="Starting audit...")
    log_output = io.StringIO()

    # Live log container — shows print output in real-time
    live_log = st.container()
    log_expander = live_log.expander("📜 Live Log (click to expand)", expanded=True)
    log_placeholder = log_expander.empty()

    # Redirect print to both streamlit and capture
    class StreamlitLogger:
        def __init__(self):
            self.logs = []
            self.original_stdout = sys.stdout

        def write(self, text):
            if text.strip():
                self.logs.append(text)
                # Update live log in UI
                try:
                    log_placeholder.code(self.get_log(), language=None)
                except Exception:
                    pass
            self.original_stdout.write(text)

        def flush(self):
            self.original_stdout.flush()

        def get_log(self):
            return "".join(self.logs)

    logger = StreamlitLogger()

    try:
        # Override interactive prompts for discover_models (streamlit can't do input())
        original_discover = audit.discover_models
        original_prompt = audit.prompt_model_selection

        def _st_discover_models():
            """Non-interactive model discovery for Streamlit."""
            if audit.JEDAI_MODEL and audit.JEDAI_MODEL.lower() != "auto":
                # User specified a model — just use it
                st.info(f"Using model: **{audit.JEDAI_MODEL}**")
                return
            # Auto mode — discover and pick first
            import requests
            base_url = audit.LOGIN_URL.rsplit('/api/', 1)[0]
            model_urls = [
                f"{base_url}/api/copilot/v1/llm/models",
                f"{base_url}/api/v1/models",
                f"{base_url}/api/copilot/v1/models",
            ]
            for url in model_urls:
                try:
                    resp = requests.get(
                        url,
                        headers={"Authorization": f"Bearer {audit.AUTH_TOKEN}",
                                 "accept": "application/json"},
                        timeout=30, verify=False,
                    )
                    if resp.status_code == 200:
                        data = resp.json()
                        models = data if isinstance(data, list) else data.get("models", data.get("data", []))
                        if models:
                            names = [m if isinstance(m, str) else m.get("id", m.get("name", str(m))) for m in models]
                            audit.AVAILABLE_MODELS = names
                            if audit.JEDAI_MODEL and audit.JEDAI_MODEL.lower() == "auto":
                                audit.JEDAI_MODEL = names[0]
                                st.info(f"Auto-selected model: **{audit.JEDAI_MODEL}**")
                            elif audit.JEDAI_MODEL in names:
                                st.info(f"Using model: **{audit.JEDAI_MODEL}**")
                            else:
                                audit.JEDAI_MODEL = names[0]
                                st.warning(f"Configured model not found. Auto-selected: **{audit.JEDAI_MODEL}**")
                            return
                except Exception:
                    continue
            if not audit.JEDAI_MODEL or audit.JEDAI_MODEL.lower() == "auto":
                st.error("Could not discover models. Please specify a model name in the sidebar.")
                st.stop()

        audit.discover_models = _st_discover_models
        audit.prompt_model_selection = lambda reason="": None  # No-op in streamlit

        # Run with logging
        sys.stdout = logger
        status = st.status("🔄 Running audit...", expanded=True)

        with status:
            # Step 0: Auth
            st.write("**[0/6]** Authenticating with JedAI...")
            progress_bar.progress(5, text="Authenticating...")
            audit.jedai_login()
            _st_discover_models()
            st.write(f"✅ Authenticated. Model: `{audit.JEDAI_MODEL}`")

            # Step 1: Parse PPTX
            st.write(f"**[1/6]** Parsing PPTX: `{pptx_file.name}`...")
            progress_bar.progress(10, text="Parsing PPTX...")
            slides = audit.parse_pptx(pptx_path)
            slide_text = "\n\n".join(
                f"[Slide {s['slide_number']}]\n{s['text']}" for s in slides if s["text"].strip()
            )
            st.write(f"✔ Extracted text from {len(slides)} slides ({len(slide_text):,} chars)")

            # Step 2: Parse manuals
            st.write(f"**[2/6]** Parsing {len(manual_paths)} manual file(s)...")
            progress_bar.progress(15, text="Parsing manuals...")
            manual_texts = []
            for mp in manual_paths:
                from pathlib import Path
                txt = audit.parse_manual(mp)
                manual_texts.append(f"--- FILE: {Path(mp).name} ---\n{txt}")
                st.write(f"  ✔ {Path(mp).name}: {len(txt):,} chars")
            combined_manual = "\n\n".join(manual_texts)

            # Step 3: Extract commands
            st.write("**[3/6]** Extracting commands/options/attributes from slides...")
            progress_bar.progress(20, text="Extracting commands...")
            regex_extracted = audit.extract_commands_regex(slides)
            st.write(f"  Regex found {len(regex_extracted)} commands")

            slide_chunks = audit.chunk_text(slide_text, audit.SLIDE_CHUNK_SIZE)
            regex_summary = "\n".join(
                f"  - {item['command']} opts:[{' '.join(item.get('options', []))}] "
                f"attrs:[{' '.join(item.get('attributes', []))}] [slides: {item['slide_numbers']}]"
                for item in regex_extracted
            )

            all_extracted = []
            from concurrent.futures import ThreadPoolExecutor, as_completed

            def _refine_chunk_st(args):
                i, s_chunk, total = args
                prompt = audit.build_extract_commands_prompt(s_chunk, regex_summary)
                raw = audit.call_jedai(prompt)
                parsed = audit.parse_jedai_response(raw)
                if isinstance(parsed, list):
                    items = parsed
                elif isinstance(parsed, dict):
                    items = parsed.get("commands", parsed.get("data", []))
                    if not items:
                        for v in parsed.values():
                            if isinstance(v, list):
                                items = v
                                break
                else:
                    items = []
                return i, items

            st.write(f"  🤖 LLM refining {len(slide_chunks)} chunk(s) in parallel (max {max_workers} workers)...")
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(_refine_chunk_st, (i, chunk, len(slide_chunks)))
                           for i, chunk in enumerate(slide_chunks, start=1)]
                for future in as_completed(futures):
                    i, items = future.result()
                    all_extracted.extend(items)

            # Merge
            seen_cmds = {}
            for item in regex_extracted:
                cmd = item.get("command", "").strip()
                if cmd:
                    seen_cmds[cmd] = item
            for item in all_extracted:
                cmd = item.get("command", "").strip()
                if not cmd or cmd.startswith("-"):
                    continue
                # For generic mode: skip single-word commands (must contain underscore)
                # For conformal mode: allow multi-word commands (space-separated)
                if product_mode.lower() != "conformal":
                    if "_" not in cmd:
                        continue
                else:
                    if " " not in cmd and "_" not in cmd:
                        continue
                if cmd in seen_cmds:
                    old_sn = str(seen_cmds[cmd].get("slide_numbers", ""))
                    new_sn = str(item.get("slide_numbers", ""))
                    merged = set(s.strip() for s in f"{old_sn},{new_sn}".split(",") if s.strip())
                    seen_cmds[cmd]["slide_numbers"] = ",".join(sorted(merged, key=lambda x: int(x) if x.isdigit() else 0))
                    if item.get("context"):
                        seen_cmds[cmd]["context"] = item["context"]
                    existing_opts = set(seen_cmds[cmd].get("options", []))
                    new_opts = set(item.get("options", []))
                    seen_cmds[cmd]["options"] = sorted(existing_opts | new_opts)
                    existing_attrs = set(seen_cmds[cmd].get("attributes", []))
                    new_attrs = set(item.get("attributes", []))
                    seen_cmds[cmd]["attributes"] = sorted(existing_attrs | new_attrs)
                else:
                    seen_cmds[cmd] = item
            extracted_commands = list(seen_cmds.values())
            st.write(f"✔ Total unique commands: **{len(extracted_commands)}**")

            # Show extracted commands in an expander
            with st.expander(f"📋 Extracted commands ({len(extracted_commands)})", expanded=False):
                for ec in extracted_commands:
                    opts = ", ".join(ec.get("options", []))
                    attrs = ", ".join(ec.get("attributes", []))
                    line = f"• **{ec.get('command', '?')}**"
                    if opts:
                        line += f"  opts: `{opts}`"
                    if attrs:
                        line += f"  attrs: `{attrs}`"
                    line += f"  _(slides: {ec.get('slide_numbers', '?')})_"
                    st.markdown(line)

            commands_names_summary = ", ".join(item.get("command", "") for item in extracted_commands)

            # Step 4: Verify commands
            manual_chunks = audit.chunk_text(combined_manual, audit.CHUNK_SIZE)
            total_cmds = len(extracted_commands)
            st.write(f"**[4/6]** Verifying {total_cmds} commands against manual...")
            progress_bar.progress(30, text="Verifying commands...")

            all_slide_audit = []
            skipped_commands = []  # Track commands not found in manual

            def _find_relevant_chunks_st(cmd_name, cmd_options, cmd_attrs):
                """Find manual chunks using multiple search strategies."""
                relevant_set = []
                seen_indices = set()

                # Strategy 1: Full command name
                search_term = cmd_name.lstrip("-.").split()[0].lower()
                for idx, m_chunk in enumerate(manual_chunks):
                    if search_term in m_chunk.lower():
                        if idx not in seen_indices:
                            relevant_set.append(m_chunk)
                            seen_indices.add(idx)

                # Strategy 2: Core part of command name
                parts = cmd_name.split("_")
                if len(parts) > 2:
                    core = "_".join(parts[1:]).lower()
                    if len(core) >= 4:
                        for idx, m_chunk in enumerate(manual_chunks):
                            if idx not in seen_indices and core in m_chunk.lower():
                                relevant_set.append(m_chunk)
                                seen_indices.add(idx)

                # Strategy 3: Search by significant options
                for opt in cmd_options:
                    opt_term = opt.lstrip("-").lower()
                    if len(opt_term) >= 5:
                        for idx, m_chunk in enumerate(manual_chunks):
                            if idx not in seen_indices and opt_term in m_chunk.lower():
                                relevant_set.append(m_chunk)
                                seen_indices.add(idx)

                # Strategy 4: Search by attribute names
                for attr in cmd_attrs:
                    attr_term = attr.lower()
                    if len(attr_term) >= 5:
                        for idx, m_chunk in enumerate(manual_chunks):
                            if idx not in seen_indices and attr_term in m_chunk.lower():
                                relevant_set.append(m_chunk)
                                seen_indices.add(idx)

                return relevant_set

            def _verify_cmd_st(args):
                cmd_idx, cmd_item = args
                cmd_name = cmd_item.get("command", "")
                cmd_options = cmd_item.get("options", [])
                cmd_attrs = cmd_item.get("attributes", [])
                cmd_context = cmd_item.get("context", "")
                cmd_slides = cmd_item.get("slide_numbers", "")
                options_str = ", ".join(cmd_options) if cmd_options else "(none)"
                attrs_str = ", ".join(cmd_attrs) if cmd_attrs else "(none)"

                relevant_chunks = _find_relevant_chunks_st(cmd_name, cmd_options, cmd_attrs)

                if not relevant_chunks:
                    skipped_commands.append(cmd_item)
                    return cmd_idx, cmd_name, []

                combined_relevant = "\n\n---\n\n".join(relevant_chunks)
                if len(combined_relevant) > audit.CHUNK_SIZE * 3:
                    combined_relevant = combined_relevant[:audit.CHUNK_SIZE * 3]

                prompt = audit.build_slide_audit_prompt(
                    cmd_name, options_str, attrs_str, cmd_context, cmd_slides, combined_relevant
                )
                raw_response = audit.call_jedai(prompt)
                result = audit.parse_jedai_response(raw_response)
                audit_items = result.get("slide_audit", [])
                return cmd_idx, cmd_name, audit_items

            st.write(f"  🚀 Verifying in parallel (max {max_workers} workers)...")
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [executor.submit(_verify_cmd_st, (i, cmd_item))
                           for i, cmd_item in enumerate(extracted_commands, start=1)]
                done_count = 0
                for future in as_completed(futures):
                    cmd_idx, cmd_name, audit_items = future.result()
                    done_count += 1
                    pct = 30 + int(60 * done_count / total_cmds)
                    progress_bar.progress(pct, text=f"Verified {done_count}/{total_cmds}: {cmd_name}")
                    if audit_items:
                        all_slide_audit.extend(audit_items)
                        for ai in audit_items:
                            st.write(f"  ⚠️ **{cmd_name}**: {ai.get('change_type', '?')} — "
                                     f"{ai.get('change_description', '')[:100]}")

            verified_count = total_cmds - len(skipped_commands)
            st.write(f"✔ Phase 1 complete: **{verified_count}/{total_cmds}** commands verified, "
                     f"**{len(all_slide_audit)}** changes found")

            if skipped_commands:
                with st.expander(f"⚠️ {len(skipped_commands)} command(s) not found in manual", expanded=False):
                    st.caption("These commands/options/attributes were not found in any provided manual section. "
                               "Consider providing additional manual files to improve coverage.")
                    for sc in skipped_commands:
                        opts = ", ".join(sc.get("options", []))
                        attrs = ", ".join(sc.get("attributes", []))
                        line = f"• **{sc.get('command', '?')}** (slides: {sc.get('slide_numbers', '?')})"
                        if opts:
                            line += f"  opts: `{opts}`"
                        if attrs:
                            line += f"  attrs: `{attrs}`"
                        st.markdown(line)

            # Step 5: Phase 2
            all_new_commands = []
            if audit.ENABLE_PHASE2:
                st.write("**[5/6]** Scanning for new commands...")
                progress_bar.progress(92, text="Phase 2: new commands...")

                def _scan_chunk_st(args):
                    i, m_chunk, total = args
                    prompt = audit.build_new_commands_prompt(commands_names_summary, m_chunk, i, total)
                    raw_response = audit.call_jedai(prompt)
                    result = audit.parse_jedai_response(raw_response)
                    return result.get("new_commands", [])

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = [executor.submit(_scan_chunk_st, (i, chunk, len(manual_chunks)))
                               for i, chunk in enumerate(manual_chunks, start=1)]
                    for future in as_completed(futures):
                        all_new_commands.extend(future.result())
                st.write(f"✔ Phase 2: {len(all_new_commands)} new commands found")
            else:
                st.write("**[5/6]** Phase 2 skipped (disabled)")

            # Step 6: Write report
            st.write("**[6/6]** Writing Excel report...")
            progress_bar.progress(95, text="Writing report...")
            audit.write_excel_report(all_slide_audit, all_new_commands, output_path)
            progress_bar.progress(100, text="✅ Audit complete!")
            status.update(label="✅ Audit complete!", state="complete")

        # Restore stdout
        elapsed = time.time() - start_time
        sys.stdout = logger.original_stdout

        # ── Format elapsed time ──
        el_h = int(elapsed // 3600)
        el_m = int((elapsed % 3600) // 60)
        el_s = int(elapsed % 60)
        time_str = f"{el_h:02d}:{el_m:02d}:{el_s:02d}"

        # ── Results ──
        st.success(f"Audit complete in **{time_str}**! **{len(extracted_commands)}** commands checked, "
                   f"**{len(all_slide_audit)}** changes found, "
                   f"**{len(all_new_commands)}** new commands.")

        # Show results in tabs
        tab1, tab2, tab3 = st.tabs(["📋 Slide Audit", "🆕 New Commands", "📥 Download"])

        with tab1:
            if all_slide_audit:
                import pandas as pd
                from collections import OrderedDict

                # Merge multiple changes for the same command into one row
                merged = OrderedDict()
                for item in all_slide_audit:
                    cmd = item.get("command", "")
                    if cmd not in merged:
                        merged[cmd] = {
                            "command": cmd,
                            "change_types": [],
                            "change_descriptions": [],
                            "slide_numbers": set(),
                            "priorities": [],
                            "actions": [],
                        }
                    desc = item.get("change_description", "").strip()
                    ctype = item.get("change_type", "").strip()
                    dedup_key = (ctype, desc[:80])
                    existing_keys = list(zip(merged[cmd]["change_types"],
                                              [d[:80] for d in merged[cmd]["change_descriptions"]]))
                    if dedup_key in existing_keys:
                        continue
                    merged[cmd]["change_types"].append(ctype)
                    merged[cmd]["change_descriptions"].append(desc)
                    merged[cmd]["priorities"].append(item.get("priority", "Medium"))
                    merged[cmd]["actions"].append(item.get("action_needed", "").strip())
                    for s in str(item.get("slide_numbers", "")).split(","):
                        s = s.strip()
                        if s:
                            merged[cmd]["slide_numbers"].add(s)

                priority_rank = {"High": 3, "Medium": 2, "Low": 1}
                merged_rows = []
                for cmd, info in merged.items():
                    if len(info["change_types"]) == 1:
                        ct = info["change_types"][0]
                        cd = info["change_descriptions"][0]
                        act = info["actions"][0]
                    else:
                        ct = "; ".join(dict.fromkeys(info["change_types"]))
                        cd = "\n".join(f"{i}. {d}" for i, d in enumerate(info["change_descriptions"], 1))
                        act = "\n".join(f"{i}. {a}" for i, a in enumerate(info["actions"], 1) if a)
                    pri = max(info["priorities"], key=lambda p: priority_rank.get(p, 0))
                    sns = ",".join(sorted(info["slide_numbers"],
                                          key=lambda x: int(x) if x.isdigit() else 0))
                    merged_rows.append({
                        "Command": cmd, "Change Type": ct, "Change Description": cd,
                        "Slide Numbers": sns, "Priority": pri, "Action Needed": act,
                    })

                df = pd.DataFrame(merged_rows)

                # Color-code by change type
                def highlight_change(row):
                    ct = str(row.get("Change Type", ""))
                    if "Command" in ct:
                        return ["background-color: #D6DCE4"] * len(row)
                    elif "Option" in ct:
                        return ["background-color: #BDD7EE"] * len(row)
                    elif "Attribute" in ct:
                        return ["background-color: #E2EFDA"] * len(row)
                    return [""] * len(row)

                styled_df = df.style.apply(highlight_change, axis=1)
                st.dataframe(styled_df, use_container_width=True)
            else:
                st.info("No changes found.")

        with tab2:
            if all_new_commands:
                import pandas as pd
                df2 = pd.DataFrame(all_new_commands)
                st.dataframe(df2, use_container_width=True)
            else:
                st.info("No new commands found (or Phase 2 was disabled).")

        with tab3:
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Excel Report",
                    data=f.read(),
                    file_name="slide_audit_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

        # Show log in expander
        with st.expander("🔍 Detailed log", expanded=False):
            st.text(logger.get_log())

    except Exception as e:
        sys.stdout = logger.original_stdout
        import traceback
        tb = traceback.format_exc()
        st.error(f"❌ Error: {e}")
        st.code(tb, language="python")
        with st.expander("🔍 Full log output", expanded=True):
            st.code(logger.get_log(), language=None)

    finally:
        sys.stdout = logger.original_stdout
        audit.discover_models = original_discover
        audit.prompt_model_selection = original_prompt

elif not can_run:
    missing = []
    if not username or not password:
        missing.append("LDAP credentials")
    if pptx_file is None:
        missing.append("PPTX file")
    if not manual_files:
        missing.append("manual file(s)")
    st.info(f"Please provide: {', '.join(missing)}")
