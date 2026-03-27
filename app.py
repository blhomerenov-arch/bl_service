elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # KPIs (basiques pour le moment)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Lignes", len(motif_df))
        col2.metric("Colonnes détectées", len(motif_df.columns))

        st.divider()
        st.subheader("📈 Statistiques des Motifs")

        # ====================== DÉTECTION AMÉLIORÉE ======================
        motif_col = find_column(motif_df, [
            'detail motif', 'détail motif', 'motif', 
            'pc mauvais', 'code motif', 'produit'
        ])

        # Si rien n'est trouvé, on prend la 6ème colonne (index 5) qui semble contenir les motifs
        if not motif_col:
            if len(motif_df.columns) > 5:
                motif_col = motif_df.columns[5]   # "PC mauvais 1221/2 P4P5"
                st.info(f"Colonne utilisée automatiquement : **{motif_col}**")

        if motif_col and motif_col in motif_df.columns:
            # Nettoyage
            motif_series = motif_df[motif_col].astype(str).str.strip()
            motif_series = motif_series[(motif_series != "") & 
                                      (motif_series != "nan") & 
                                      (motif_series != "None") & 
                                      (motif_series != "nan")]

            motif_count = motif_series.value_counts().head(15)

            if not motif_count.empty:
                # Graphique en barres
                st.subheader("📊 Top 15 des Motifs")
                fig_bar = px.bar(
                    x=motif_count.index,
                    y=motif_count.values,
                    title=f"Top 15 Motifs - Colonne : {motif_col}",
                    labels={"x": "Motif", "y": "Nombre"},
                    text=motif_count.values
                )
                fig_bar.update_layout(xaxis_tickangle=-45, height=550, margin=dict(b=200))
                st.plotly_chart(fig_bar, use_container_width=True)

                # Graphique en cercle
                st.subheader("🥧 Répartition des Motifs")
                top10 = motif_series.value_counts().head(10)
                fig_pie = px.pie(
                    names=top10.index, 
                    values=top10.values, 
                    title="Répartition en pourcentage"
                )
                fig_pie.update_traces(textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

                # Tableau
                st.dataframe(
                    motif_count.reset_index().rename(columns={"index": "Motif", "count": "Nombre"}),
                    use_container_width=True
                )
            else:
                st.warning("La colonne existe mais ne contient aucun motif pour le moment.")
                st.info("Remplis la colonne des motifs dans ton fichier Excel.")
        else:
            st.error("Impossible de détecter la colonne des motifs.")
            st.write("**Colonnes disponibles :**")
            st.write(list(motif_df.columns))

    except Exception as e:
        st.error(f"Erreur : {str(e)}")
