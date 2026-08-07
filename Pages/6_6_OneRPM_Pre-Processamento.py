import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.title("Processamento de Royalties")
st.write("Faça upload dos arquivos xlsx contendo as planilhas Masters, Youtube Channels e Shares In & Out")

st.divider()

# Upload de múltiplos arquivos
uploaded_files = st.file_uploader("Selecione os arquivos xlsx", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    try:
        # Inicializar dataframes vazios para consolidação
        all_masters = []
        all_youtube = []
        all_shares = []
        
        # Processar cada arquivo
        st.subheader("Arquivos carregados")
        for i, uploaded_file in enumerate(uploaded_files, 1):
            st.write(f"{i}. {uploaded_file.name}")
            
            # Ler as planilhas de cada arquivo
            df_masters = pd.read_excel(uploaded_file, sheet_name='Masters')
            df_youtube = pd.read_excel(uploaded_file, sheet_name='Youtube Channels')
            df_shares = pd.read_excel(uploaded_file, sheet_name='Shares In & Out')
            
            # Adicionar às listas
            all_masters.append(df_masters)
            all_youtube.append(df_youtube)
            all_shares.append(df_shares)
        
        # Consolidar todos os dataframes
        df_masters = pd.concat(all_masters, ignore_index=True)
        df_youtube = pd.concat(all_youtube, ignore_index=True)
        df_shares = pd.concat(all_shares, ignore_index=True)
        
        st.success(f"{len(uploaded_files)} arquivo(s) carregado(s) e consolidado(s) com sucesso")
        st.divider()
        
        # RESUMO 1: Soma por planilha e moeda
        st.subheader("Valores por planilha original (consolidado)")
        
        def show_summary_df(df, sheet_name):
            st.write(f"**{sheet_name}:**")
            if df.empty or 'Net' not in df.columns or df['Net'].isna().all() or df['Net'].sum() == 0:
                st.write("*Sem rendimentos*")
            else:
                if 'Currency' in df.columns:
                    summary = df.groupby('Currency')['Net'].sum().reset_index()
                    summary.columns = ['Moeda', 'Valor']
                    st.dataframe(summary, hide_index=True, use_container_width=True)
                else:
                    st.write(f"Total: {df['Net'].sum():,.2f}")
        
        show_summary_df(df_masters, "Masters")
        show_summary_df(df_youtube, "Youtube Channels")
        show_summary_df(df_shares, "Shares In & Out")
        
        # Análise Share-in e Share-out
        if not df_shares.empty and 'Share Type' in df_shares.columns and 'Net' in df_shares.columns:
            st.write("")
            
            share_in = df_shares[df_shares['Share Type'] == 'In']
            share_out = df_shares[df_shares['Share Type'] == 'Out']
            
            st.write("**Share-In:**")
            if share_in.empty or share_in['Net'].sum() == 0:
                st.write("*Sem rendimentos*")
            else:
                if 'Currency' in share_in.columns:
                    summary_in = share_in.groupby('Currency')['Net'].sum().reset_index()
                    summary_in.columns = ['Moeda', 'Valor']
                    st.dataframe(summary_in, hide_index=True, use_container_width=True)
            
            st.write("**Share-Out:**")
            if share_out.empty or share_out['Net'].sum() == 0:
                st.write("*Sem rendimentos*")
            else:
                if 'Currency' in share_out.columns:
                    summary_out = share_out.groupby('Currency')['Net'].sum().reset_index()
                    summary_out.columns = ['Moeda', 'Valor']
                    st.dataframe(summary_out, hide_index=True, use_container_width=True)
        
        st.divider()
        
        # Processamento dos dados
        st.subheader("Processamento dos dados")
        
        # Filtrar Shares In & Out
        # Excluir "listener-1703345420400" da coluna Receiver Name
        df_shares_filtered = df_shares[df_shares['Receiver Name'] != 'listener-1703345420400'].copy()
        
        # Separar YouTube Video para concatenar com Youtube Channels
        df_shares_youtube = df_shares_filtered[df_shares_filtered['Product Type'] == 'YouTube Video'].copy()
        
        # Excluir YouTube Video do restante (para concatenar com Masters)
        df_shares_masters = df_shares_filtered[df_shares_filtered['Product Type'] != 'YouTube Video'].copy()
        
        # Mapeamento e concatenação Masters + Shares In & Out
        df_shares_masters_mapped = df_shares_masters.rename(columns={
            'Title': 'Track Title',
            'Artists': 'Artists',
            'Product Type': 'Product Type',
            'ID': 'ISRC',
            'Parent ID': 'UPC',
            'Store': 'Store',
            'Territory': 'Territory',
            'Sale Type': 'Sale Type',
            'Transaction Month': 'Transaction Month',
            'Accounted Date': 'Accounted Date',
            'Currency': 'Currency',
            'Quantity': 'Quantity',
            'Net': 'Net'
        })
        
        # Manter apenas as colunas mapeadas
        columns_masters = ['Track Title', 'Artists', 'Product Type', 'ISRC', 'UPC', 'Store', 
                          'Territory', 'Sale Type', 'Transaction Month', 'Accounted Date', 
                          'Currency', 'Quantity', 'Net']
        
        df_shares_masters_mapped = df_shares_masters_mapped[columns_masters]
        df_masters_concat = pd.concat([df_masters, df_shares_masters_mapped], ignore_index=True)
        
        # Mapeamento e concatenação Youtube Channels + Shares In & Out (YouTube Video)
        df_shares_youtube_mapped = df_shares_youtube.rename(columns={
            'Title': 'Video Title',
            'ID': 'Video ID',
            'Parent ID': 'Channel ID',
            'Store': 'Store',
            'Territory': 'Territory',
            'Sale Type': 'Sale Type',
            'Transaction Month': 'Transaction Month',
            'Accounted Date': 'Accounted Date',
            'Currency': 'Currency',
            'Quantity': 'Quantity',
            'Net': 'Net'
        })
        
        # Manter apenas as colunas mapeadas
        columns_youtube = ['Video Title', 'Video ID', 'Channel ID', 'Store', 'Territory', 
                          'Sale Type', 'Transaction Month', 'Accounted Date', 'Currency', 
                          'Quantity', 'Net']
        
        df_shares_youtube_mapped = df_shares_youtube_mapped[columns_youtube]
        df_youtube_concat = pd.concat([df_youtube, df_shares_youtube_mapped], ignore_index=True)
        
        st.success("Dados concatenados com sucesso")
        st.divider()
        
        # RESUMO 2: Valores após concatenação
        st.subheader("Valores após concatenação")
        
        st.write("**Masters + Shares In & Out:**")
        if df_masters_concat.empty or df_masters_concat['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_masters = df_masters_concat.groupby('Currency')['Net'].sum().reset_index()
            summary_masters.columns = ['Moeda', 'Valor']
            st.dataframe(summary_masters, hide_index=True, use_container_width=True)
        
        st.write("**Youtube Channels:**")
        if df_youtube_concat.empty or df_youtube_concat['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_youtube = df_youtube_concat.groupby('Currency')['Net'].sum().reset_index()
            summary_youtube.columns = ['Moeda', 'Valor']
            st.dataframe(summary_youtube, hide_index=True, use_container_width=True)
        
        st.divider()
        
        # Inputs para taxas bancárias
        st.subheader("Taxas bancárias")
        
        col1, col2 = st.columns(2)
        with col1:
            taxa_brl = st.number_input("Taxa BRL", min_value=0.0, value=0.49, step=0.01, format="%.2f", help="Valor da taxa bancária a ser descontada proporcionalmente")
        with col2:
            taxa_usd = st.number_input("Taxa USD", min_value=0.0, value=26.00, step=0.01, format="%.2f", help="Valor da taxa bancária a ser descontada proporcionalmente")
        
        st.divider()
        
        # Aplicar descontos proporcionais entre Masters e Youtube
        def apply_proportional_discount(df_masters, df_youtube, currency, discount_amount):
            if discount_amount == 0:
                return df_masters, df_youtube
            
            # Calcular totais por dataframe
            total_masters = df_masters[df_masters['Currency'] == currency]['Net'].sum()
            total_youtube = df_youtube[df_youtube['Currency'] == currency]['Net'].sum()
            total_combined = total_masters + total_youtube
            
            if total_combined == 0:
                return df_masters, df_youtube
            
            # Calcular proporção de desconto para cada dataframe
            discount_masters = discount_amount * (total_masters / total_combined)
            discount_youtube = discount_amount * (total_youtube / total_combined)
            
            # Aplicar desconto proporcional
            if total_masters > 0:
                df_masters.loc[df_masters['Currency'] == currency, 'Net'] = \
                    df_masters.loc[df_masters['Currency'] == currency, 'Net'] - \
                    (df_masters.loc[df_masters['Currency'] == currency, 'Net'] / total_masters * discount_masters)
            
            if total_youtube > 0:
                df_youtube.loc[df_youtube['Currency'] == currency, 'Net'] = \
                    df_youtube.loc[df_youtube['Currency'] == currency, 'Net'] - \
                    (df_youtube.loc[df_youtube['Currency'] == currency, 'Net'] / total_youtube * discount_youtube)
            
            return df_masters, df_youtube
        
        # Criar cópias para aplicar os descontos
        df_masters_final = df_masters_concat.copy()
        df_youtube_final = df_youtube_concat.copy()
        
        # Aplicar descontos de forma proporcional
        if taxa_brl > 0:
            df_masters_final, df_youtube_final = apply_proportional_discount(df_masters_final, df_youtube_final, 'BRL', taxa_brl)
        
        if taxa_usd > 0:
            df_masters_final, df_youtube_final = apply_proportional_discount(df_masters_final, df_youtube_final, 'USD', taxa_usd)
        
        # RESUMO 3: Valores após descontos
        st.subheader("Valores após descontos das taxas")
        
        st.write("**Masters + Shares In & Out:**")
        if df_masters_final.empty or df_masters_final['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_masters_final = df_masters_final.groupby('Currency')['Net'].sum().reset_index()
            summary_masters_final.columns = ['Moeda', 'Valor']
            st.dataframe(summary_masters_final, hide_index=True, use_container_width=True)
        
        st.write("**Youtube Channels:**")
        if df_youtube_final.empty or df_youtube_final['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_youtube_final = df_youtube_final.groupby('Currency')['Net'].sum().reset_index()
            summary_youtube_final.columns = ['Moeda', 'Valor']
            st.dataframe(summary_youtube_final, hide_index=True, use_container_width=True)
        
        st.divider()
        
        # Downloads
        st.subheader("Download dos resultados finais")
        
        # Função para criar arquivo Excel
        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()
        
        col1, col2 = st.columns(2)
        
        # Downloads Masters
        with col1:
            st.write("**Masters + Shares In & Out:**")
            if not df_masters_final.empty:
                # Download completo (todas as moedas)
                excel_data_all = to_excel(df_masters_final)
                st.download_button(
                    label="📥 Download Masters (Todas as moedas)",
                    data=excel_data_all,
                    file_name=f"Masters_COMPLETO_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.write("")
                st.write("*Por moeda:*")
                
                # Downloads individuais por moeda
                currencies_masters = sorted(df_masters_final['Currency'].unique())
                for currency in currencies_masters:
                    df_download = df_masters_final[df_masters_final['Currency'] == currency]
                    excel_data = to_excel(df_download)
                    st.download_button(
                        label=f"Download Masters {currency}",
                        data=excel_data,
                        file_name=f"Masters_{currency}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        # Downloads Youtube
        with col2:
            st.write("**Youtube Channels:**")
            if not df_youtube_final.empty:
                # Download completo (todas as moedas)
                excel_data_all = to_excel(df_youtube_final)
                st.download_button(
                    label="📥 Download Youtube (Todas as moedas)",
                    data=excel_data_all,
                    file_name=f"Youtube_COMPLETO_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.write("")
                st.write("*Por moeda:*")
                
                # Downloads individuais por moeda
                currencies_youtube = sorted(df_youtube_final['Currency'].unique())
                for currency in currencies_youtube:
                    df_download = df_youtube_final[df_youtube_final['Currency'] == currency]
                    excel_data = to_excel(df_download)
                    st.download_button(
                        label=f"Download Youtube {currency}",
                        data=excel_data,
                        file_name=f"Youtube_{currency}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
        st.divider()
        
        # RESUMO 4: Valores finais com totais
        st.subheader("Valores finais por moeda e dataframe")
        
        st.write("**Masters + Shares In & Out:**")
        if df_masters_final.empty or df_masters_final['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_masters_final = df_masters_final.groupby('Currency')['Net'].sum().reset_index()
            summary_masters_final.columns = ['Moeda', 'Valor']
            st.dataframe(summary_masters_final, hide_index=True, use_container_width=True)
        
        st.write("**Youtube Channels:**")
        if df_youtube_final.empty or df_youtube_final['Net'].sum() == 0:
            st.write("*Sem rendimentos*")
        else:
            summary_youtube_final = df_youtube_final.groupby('Currency')['Net'].sum().reset_index()
            summary_youtube_final.columns = ['Moeda', 'Valor']
            st.dataframe(summary_youtube_final, hide_index=True, use_container_width=True)
        
        # TOTAL GERAL: Masters + Youtube por moeda
        st.write("**TOTAL GERAL (Masters + Shares In & Out + Youtube Channels):**")
        
        # Obter todas as moedas únicas
        all_currencies = set(df_masters_final['Currency'].unique()) | set(df_youtube_final['Currency'].unique())
        
        total_data = []
        for currency in sorted(all_currencies):
            total_masters_currency = df_masters_final[df_masters_final['Currency'] == currency]['Net'].sum()
            total_youtube_currency = df_youtube_final[df_youtube_final['Currency'] == currency]['Net'].sum()
            total_currency = total_masters_currency + total_youtube_currency
            total_data.append({'Moeda': currency, 'Valor': total_currency})
        
        summary_total = pd.DataFrame(total_data)
        st.dataframe(summary_total, hide_index=True, use_container_width=True)
        
    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {str(e)}")
        st.write("Certifique-se de que todos os arquivos contêm as planilhas Masters, Youtube Channels e Shares In & Out")

else:
    st.info("Aguardando upload dos arquivos")