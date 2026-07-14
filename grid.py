"""Shared editable grid with a search filter and a reset (undo-all) button.

Streamlit's data_editor has no built-in filter/undo; this keeps a master copy
in session_state so edits survive filtering, and reset throws every edit away.
The caller's `key` must fingerprint the dataset (RO/masa/size) so a new upload
starts clean.
"""
import pandas as pd
import streamlit as st


def editable_grid(df: pd.DataFrame, key: str, column_config=None,
                  height: int = 340, allow_add: bool = True) -> pd.DataFrame:
    store, nonce_key = f"{key}__store", f"{key}__nonce"
    nonce = st.session_state.setdefault(nonce_key, 0)
    if store not in st.session_state:
        st.session_state[store] = df.copy()
    master = st.session_state[store]

    c1, c2 = st.columns([5, 1])
    q = c1.text_input("Filter", key=f"{key}__q_{nonce}",
                      placeholder="🔍 Filter — ketik untuk cari di semua kolom…",
                      label_visibility="collapsed")
    if c2.button("🔄 Reset", key=f"{key}__reset_{nonce}", use_container_width=True,
                 help="Buang semua edit & baris tambahan — kembali ke hasil olahan."):
        st.session_state[store] = df.copy()
        st.session_state[nonce_key] = nonce + 1
        st.rerun()

    if q:
        mask = master.apply(
            lambda row: row.astype(str).str.contains(q, case=False, na=False).any(),
            axis=1)
        view = master[mask]
        edited = st.data_editor(
            view, num_rows="fixed", use_container_width=True, height=height,
            key=f"{key}__ed_{nonce}_{q}", column_config=column_config or {})
        master.loc[edited.index] = edited          # merge edits back into master
        st.caption(f"{len(view)} dari {len(master)} baris cocok — edit tetap "
                   "tersimpan; hapus filter untuk lihat semua / tambah baris.")
    else:
        edited = st.data_editor(
            master, num_rows="dynamic" if allow_add else "fixed",
            use_container_width=True, height=height,
            key=f"{key}__ed_{nonce}", column_config=column_config or {})
        st.session_state[store] = edited
        master = edited
    return st.session_state[store]
