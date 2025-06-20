import io
import os
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Dismac: Reserva de Entrega de Mercadería", layout="wide")

# ─────────────────────────────────────────────────────────────
# 1. Configuration
# ─────────────────────────────────────────────────────────────
try:
    SITE_URL = os.getenv("SP_SITE_URL") or st.secrets["SP_SITE_URL"]
    FILE_ID = os.getenv("SP_FILE_ID") or st.secrets["SP_FILE_ID"]
    USERNAME = os.getenv("SP_USERNAME") or st.secrets["SP_USERNAME"]
    PASSWORD = os.getenv("SP_PASSWORD") or st.secrets["SP_PASSWORD"]
    
    # Email configuration
    EMAIL_HOST = os.getenv("EMAIL_HOST") or st.secrets["EMAIL_HOST"]
    EMAIL_PORT = int(os.getenv("EMAIL_PORT") or st.secrets["EMAIL_PORT"])
    EMAIL_USER = os.getenv("EMAIL_USER") or st.secrets["EMAIL_USER"]
    EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") or st.secrets["EMAIL_PASSWORD"]
    
except KeyError as e:
    st.error(f"🔒 Falta configuración: {e}")
    st.stop()

# ─────────────────────────────────────────────────────────────
# 2. Excel Download Functions - UPDATED TO INCLUDE GESTION SHEET
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=60)  # Reduced cache time to 1 minute
def download_excel_to_memory():
    """Download Excel file from SharePoint to memory - INCLUDES ALL SHEETS"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        if file is None:
            raise Exception("File object is None - FILE_ID may be incorrect")
            
        ctx.load(file)
        ctx.execute_query()
        
        # Download to memory
        file_content = io.BytesIO()
        
        # Try multiple download methods
        try:
            file.download(file_content)
            ctx.execute_query()
        except TypeError as e:
            try:
                response = file.download()
                if response is None:
                    raise Exception("Download response is None")
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
            except Exception as e2:
                try:
                    file.download_session(file_content)
                    ctx.execute_query()
                except Exception as e3:
                    raise Exception(f"All download methods failed: {e}, {e2}, {e3}")
        
        file_content.seek(0)
        
        # Load all sheets - UPDATED
        credentials_df = pd.read_excel(file_content, sheet_name="proveedor_credencial", dtype=str)
        reservas_df = pd.read_excel(file_content, sheet_name="proveedor_reservas")
        
        # Try to load gestion sheet, create empty if doesn't exist - NEW
        try:
            gestion_df = pd.read_excel(file_content, sheet_name="proveedor_gestion")
        except ValueError:
            # Create empty gestion dataframe with required columns if sheet doesn't exist
            gestion_df = pd.DataFrame(columns=[
                'Orden_de_compra', 'Proveedor', 'Numero_de_bultos',
                'Hora_llegada', 'Hora_inicio_atencion', 'Hora_fin_atencion',
                'Tiempo_espera', 'Tiempo_atencion', 'Tiempo_total', 'Tiempo_retraso',
                'numero_de_semana', 'hora_de_reserva'
            ])
        
        return credentials_df, reservas_df, gestion_df
        
    except Exception as e:
        st.error(f"Error descargando Excel: {str(e)}")
        st.error(f"SITE_URL: {SITE_URL}")
        st.error(f"FILE_ID: {FILE_ID}")
        st.error(f"Error type: {type(e).__name__}")
        return None, None, None

def save_booking_to_excel(new_booking):
    """Save new booking to Excel file - PRESERVES ALL SHEETS"""
    try:
        # Clear cache before loading to get fresh data
        download_excel_to_memory.clear()
        
        # Load current data - UPDATED TO LOAD ALL SHEETS
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
            st.error("❌ No se pudo cargar el archivo Excel")
            return False
        
        # Add new booking
        new_row = pd.DataFrame([new_booking])
        updated_reservas_df = pd.concat([reservas_df, new_row], ignore_index=True)
        
        # Authenticate and upload
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Create Excel file - UPDATED TO SAVE ALL SHEETS
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            credentials_df.to_excel(writer, sheet_name="proveedor_credencial", index=False)
            updated_reservas_df.to_excel(writer, sheet_name="proveedor_reservas", index=False)
            gestion_df.to_excel(writer, sheet_name="proveedor_gestion", index=False)  # NEW - PRESERVE GESTION SHEET
        
        # Get the file info
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        file_name = file.properties['Name']
        server_relative_url = file.properties['ServerRelativeUrl']
        folder_url = server_relative_url.replace('/' + file_name, '')
        
        # Upload the updated file
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        excel_buffer.seek(0)
        folder.files.add(file_name, excel_buffer.getvalue(), True)
        ctx.execute_query()
        
        # Clear cache after successful save
        download_excel_to_memory.clear()
        
        return True
        
    except Exception as e:
        st.error(f"❌ Error guardando reserva: {str(e)}")
        # Clear cache even on failure to prevent stale data
        download_excel_to_memory.clear()
        return False

# ─────────────────────────────────────────────────────────────
# 3. Email Functions
# ─────────────────────────────────────────────────────────────
def download_pdf_attachment():
    """Download PDF attachment from SharePoint"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Target filename and exact path
        target_filename = "GUIA_DEL_SELLER_DISMAC_MARKETPLACE_Rev._1.pdf"
        file_path = f"/personal/ljbyon_dismac_com_bo/Documents/{target_filename}"
        
        try:
            # Try to get the file directly
            pdf_file = ctx.web.get_file_by_server_relative_url(file_path)
            ctx.load(pdf_file)
            ctx.execute_query()
            
        except Exception as e:
            # Fallback: List files in Documents folder
            try:
                folder = ctx.web.get_folder_by_server_relative_url("/personal/ljbyon_dismac_com_bo/Documents")
                files = folder.files
                ctx.load(files)
                ctx.execute_query()
                
                found_files = []
                pdf_file = None
                
                for file in files:
                    filename = file.name
                    found_files.append(filename)
                    
                    # Check if this is our target file
                    if filename == target_filename:
                        pdf_file = file
                        break
                
                # If still not found, try any PDF
                if pdf_file is None:
                    pdf_files = [f for f in found_files if f.lower().endswith('.pdf')]
                    
                    if pdf_files:
                        # Use the first PDF found
                        first_pdf = pdf_files[0]
                        pdf_file_path = f"/personal/ljbyon_dismac_com_bo/Documents/{first_pdf}"
                        pdf_file = ctx.web.get_file_by_server_relative_url(pdf_file_path)
                        ctx.load(pdf_file)
                        ctx.execute_query()
                    else:
                        raise Exception(f"No se encontró {target_filename} ni otros PDFs en Documents")
                        
            except Exception as e2:
                raise Exception(f"No se pudo acceder a Documents: {str(e2)}")
        
        if pdf_file is None:
            raise Exception("No se pudo cargar el archivo PDF")
        
        # Download PDF to memory
        pdf_content = io.BytesIO()
        
        try:
            pdf_file.download(pdf_content)
            ctx.execute_query()
        except TypeError:
            try:
                response = pdf_file.download()
                ctx.execute_query()
                pdf_content = io.BytesIO(response.content)
            except:
                pdf_file.download_session(pdf_content)
                ctx.execute_query()
        
        pdf_content.seek(0)
        pdf_data = pdf_content.getvalue()
        
        # Get filename
        try:
            filename = pdf_file.properties.get('Name', target_filename)
        except:
            filename = target_filename
        
        return pdf_data, filename
        
    except Exception as e:
        # Only show error if PDF download fails
        st.warning(f"No se pudo descargar el archivo adjunto: {str(e)}")
        return None, None

def send_booking_email(supplier_email, supplier_name, booking_details, cc_emails=None):
    """Send booking confirmation email with PDF attachment"""
    try:
        # Use provided CC emails or default
        if cc_emails is None or len(cc_emails) == 0:
            cc_emails = ["leonardo.byon@gmail.com"]
        else:
            # Add default email to the CC list if not already present
            if "leonardo.byon@gmail.com" not in cc_emails:
                cc_emails = cc_emails + ["leonardo.byon@gmail.com"]
        
        # Email content
        subject = "Confirmación de Reserva para Entrega de Mercadería"
        
        # Format dates for email display
        display_fecha = booking_details['Fecha'].split(' ')[0]  # Remove time part for display
        display_hora = booking_details['Hora'].rsplit(':', 1)[0]  # Remove seconds for display
        
        body = f"""
        Hola {supplier_name},
        
        Su reserva de entrega ha sido confirmada exitosamente.
        
        DETALLES DE LA RESERVA:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        📅 Fecha: {display_fecha}
        🕐 Horario: {display_hora}
        📦 Número de bultos: {booking_details['Numero_de_bultos']}
        📋 Orden de compra: {booking_details['Orden_de_compra']}
        
        INSTRUCCIONES:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        • Respeta el horario reservado para tu entrega.
        • En caso de retraso, podrías tener que esperar hasta el próximo cupo disponible del día o reprogramar tu entrega.
        • Dismac no se responsabiliza por los tiempos de espera ocasionados por llegadas fuera de horario.
        • Además, según el tipo de venta, es importante considerar lo siguiente:
          - Venta al contado: Debes entregar el pedido junto con la factura a nombre del comprador y tres (3) copias de la orden de compra.
          - Venta en minicuotas: Debes entregar el pedido junto con la factura a nombre de Dismatec S.A. y una (1) copia de la orden de compra.
        
        📎 Se adjunta documento con instrucciones adicionales.
        
        REQUISITOS DE SEGURIDAD
        • Pantalón largo, sin rasgados
        • Botines de seguridad
        • Casco de seguridad
        • Chaleco o camisa con reflectivo
        • No está permitido manillas, cadenas, y principalmente masticar coca.

        Gracias por utilizar nuestro sistema de reservas.
        
        Saludos cordiales,
        Equipo de Almacén Dismac
        """
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = supplier_email
        msg['Cc'] = ', '.join(cc_emails)
        msg['Subject'] = subject
        
        # Add body
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Download and attach PDF
        pdf_data, pdf_filename = download_pdf_attachment()
        if pdf_data:
            attachment = MIMEBase('application', 'octet-stream')
            attachment.set_payload(pdf_data)
            encoders.encode_base64(attachment)
            attachment.add_header(
                'Content-Disposition',
                f'attachment; filename= {pdf_filename}'
            )
            msg.attach(attachment)
        
        # Send email
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        
        # Send to supplier + CC recipients
        all_recipients = [supplier_email] + cc_emails
        text = msg.as_string()
        server.sendmail(EMAIL_USER, all_recipients, text)
        server.quit()
        
        return True, cc_emails
        
    except Exception as e:
        st.error(f"Error enviando email: {str(e)}")
        return False, []

# ─────────────────────────────────────────────────────────────
# 4. Time Slot Functions
# ─────────────────────────────────────────────────────────────
def generate_time_slots():
    """Generate available time slots - showing start time only"""
    # Monday-Friday: 9:00-16:00, Saturday: 9:00-12:00
    weekday_slots = []
    saturday_slots = []
    
    # Weekday slots (9:00-16:00)
    start_hour = 9
    end_hour = 16
    for hour in range(start_hour, end_hour):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            weekday_slots.append(start_time)
    
    # Saturday slots (9:00-12:00)
    for hour in range(9, 12):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            saturday_slots.append(start_time)
    
    return weekday_slots, saturday_slots

def get_available_slots(selected_date, reservas_df):
    """Get available slots for a date"""
    weekday_slots, saturday_slots = generate_time_slots()
    
    # Sunday = 6, no work
    if selected_date.weekday() == 6:
        return []
    
    # Saturday = 5
    if selected_date.weekday() == 5:
        all_slots = saturday_slots
    else:
        all_slots = weekday_slots
    
    # Filter booked slots
    date_str = selected_date.strftime('%Y-%m-%d')
    booked_slots = reservas_df[reservas_df['Fecha'] == date_str]['Hora'].tolist()
    
    return [slot for slot in all_slots if slot not in booked_slots]

# ─────────────────────────────────────────────────────────────
# 5. Authentication Function - UPDATED TO USE ALL SHEETS
# ─────────────────────────────────────────────────────────────
def authenticate_user(usuario, password):
    """Authenticate user against Excel data and get email + CC emails"""
    credentials_df, _, _ = download_excel_to_memory()  # UPDATED - Now returns 3 values
    
    if credentials_df is None:
        return False, "Error al cargar credenciales", None, None
    
    # Clean and compare (all data is already strings)
    df_usuarios = credentials_df['usuario'].str.strip()
    
    input_usuario = str(usuario).strip()
    input_password = str(password).strip()
    
    # Find user row
    user_row = credentials_df[df_usuarios == input_usuario]
    if user_row.empty:
        return False, "Usuario no encontrado", None, None
    
    # Get stored password and clean it
    stored_password = str(user_row.iloc[0]['password']).strip()
    
    # Compare passwords
    if stored_password == input_password:
        # Get email
        email = None
        try:
            email = user_row.iloc[0]['Email']
            if str(email) == 'nan' or email is None:
                email = None
        except:
            email = None
        
        # Get CC emails
        cc_emails = []
        try:
            cc_data = user_row.iloc[0]['cc']
            if str(cc_data) != 'nan' and cc_data is not None and str(cc_data).strip():
                # Parse semicolon-separated emails
                cc_emails = [email.strip() for email in str(cc_data).split(';') if email.strip()]
        except Exception as e:
            cc_emails = []
        
        return True, "Autenticación exitosa", email, cc_emails
    
    return False, "Contraseña incorrecta", None, None

# ─────────────────────────────────────────────────────────────
# 6. Main App - UPDATED TO USE ALL SHEETS
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚚 Dismac: Reserva de Entrega de Mercadería")
    
    # Force refresh button for debugging
    col1, col2, col3 = st.columns([1, 1, 3])
    with col1:
        if st.button("🔄 Actualizar Datos"):
            download_excel_to_memory.clear()
            st.success("Cache limpiado")
            st.rerun()
    
    # Download Excel when app starts - UPDATED
    with st.spinner("Cargando datos..."):
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()  # UPDATED - Now gets 3 values
    
    if credentials_df is None:
        st.error("❌ Error al cargar archivo")
        return
    
    # Debug info (remove after testing)
    with st.expander("🔍 Debug Info"):
        st.write(f"📊 Total reservas en Excel: {len(reservas_df)}")
        if len(reservas_df) > 0:
            st.write("📅 Últimas 3 reservas:")
            st.dataframe(reservas_df.tail(3)[['Fecha', 'Hora', 'Proveedor']])
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    if 'supplier_email' not in st.session_state:
        st.session_state.supplier_email = None
    if 'supplier_cc_emails' not in st.session_state:
        st.session_state.supplier_cc_emails = []
    
    # Authentication
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar Sesión")
            
            if submitted:
                if usuario and password:
                    is_valid, message, email, cc_emails = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = usuario
                        st.session_state.supplier_email = email
                        st.session_state.supplier_cc_emails = cc_emails
                        # Clear any previous session data
                        st.session_state.orden_compra_list = ['']
                        if 'selected_slot' in st.session_state:
                            del st.session_state.selected_slot
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.warning("Complete todos los campos")
    
    # Booking interface
    else:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader(f"Bienvenido, {st.session_state.supplier_name}")
        with col2:
            if st.button("Cerrar Sesión"):
                st.session_state.authenticated = False
                st.session_state.supplier_name = None
                st.session_state.supplier_email = None
                st.session_state.supplier_cc_emails = []
                # Clear booking session data
                st.session_state.orden_compra_list = ['']
                if 'selected_slot' in st.session_state:
                    del st.session_state.selected_slot
                st.rerun()
        
        st.markdown("---")
        
        # Date selection
        st.subheader("📅 Seleccionar Fecha")
        st.markdown('<p style="color: red; font-size: 14px; margin-top: -10px;">Le rogamos seleccionar la fecha y el horario con atención, ya que, una vez confirmados, no podrán ser modificados ni cancelados.</p>', unsafe_allow_html=True)
        today = datetime.now().date()
        max_date = today + timedelta(days=30)
        
        selected_date = st.date_input(
            "Fecha de entrega",
            min_value=today,
            max_value=max_date,
            value=today
        )
        
        # Check if Sunday
        if selected_date.weekday() == 6:
            st.warning("⚠️ No trabajamos los domingos")
            return
        
        # Time slot selection
        st.subheader("🕐 Horarios Disponibles")
        
        # Force fresh data for slot availability
        download_excel_to_memory.clear()
        _, fresh_reservas_df, _ = download_excel_to_memory()
        
        # Generate all slots and check availability
        weekday_slots, saturday_slots = generate_time_slots()
        
        if selected_date.weekday() == 5:  # Saturday
            all_slots = saturday_slots
        else:  # Monday-Friday
            all_slots = weekday_slots
        
        # Get booked slots for this date - FIXED: Handle new date/time format
        date_str = selected_date.strftime('%Y-%m-%d') + ' 00:00:00'  # Match the new format we save
        booked_reservas = fresh_reservas_df[fresh_reservas_df['Fecha'] == date_str]['Hora'].tolist()
        
        # Debug: Show what we found in Excel for this date
        with st.expander(f"🔍 Reservas para {selected_date}"):
            date_reservas = fresh_reservas_df[fresh_reservas_df['Fecha'] == date_str]
            if not date_reservas.empty:
                st.dataframe(date_reservas[['Hora', 'Proveedor', 'Orden_de_compra']])
            else:
                st.write("No hay reservas para esta fecha")
        
        # Convert booked slots to "09:00" format for comparison
        booked_slots = []
        for booked_hora in booked_reservas:
            if ':' in str(booked_hora):
                # Handle both old "9:00:00" and new "09:00:00" formats
                parts = str(booked_hora).split(':')
                formatted_slot = f"{int(parts[0]):02d}:{parts[1]}"
                booked_slots.append(formatted_slot)
            else:
                booked_slots.append(str(booked_hora))
        
        if not all_slots:
            st.warning("❌ No hay horarios para esta fecha")
            return
        
        # Display slots (2 per row)
        selected_slot = None
        
        for i in range(0, len(all_slots), 2):
            col1, col2 = st.columns(2)
            
            # First slot
            slot1 = all_slots[i]
            is_booked1 = slot1 in booked_slots
            
            with col1:
                if is_booked1:
                    st.button(f"🚫 {slot1} (Ocupado)", disabled=True, key=f"slot_{i}", use_container_width=True)
                else:
                    if st.button(f"✅ {slot1}", key=f"slot_{i}", use_container_width=True):
                        selected_slot = slot1
            
            # Second slot (if exists)
            if i + 1 < len(all_slots):
                slot2 = all_slots[i + 1]
                is_booked2 = slot2 in booked_slots
                
                with col2:
                    if is_booked2:
                        st.button(f"🚫 {slot2} (Ocupado)", disabled=True, key=f"slot_{i+1}", use_container_width=True)
                    else:
                        if st.button(f"✅ {slot2}", key=f"slot_{i+1}", use_container_width=True):
                            selected_slot = slot2
        
        # Booking form with MULTIPLE ORDEN DE COMPRA
        if selected_slot or 'selected_slot' in st.session_state:
            if selected_slot:
                st.session_state.selected_slot = selected_slot
            
            st.markdown("---")
            st.subheader("📦 Información de Entrega")
            
            # Initialize orden de compra list in session state - reset for each booking session
            if 'orden_compra_list' not in st.session_state or not st.session_state.orden_compra_list:
                st.session_state.orden_compra_list = ['']
            
            # Date and time info (outside form so it doesn't reset)
            st.info(f"📅 Fecha: {selected_date}")
            st.info(f"🕐 Horario: {st.session_state.selected_slot}")
            
            # Number of bultos (outside form)
            numero_bultos = st.number_input(
                "📦 Número de bultos", 
                min_value=1, 
                value=1,
                help="Cantidad de bultos o paquetes a entregar"
            )
            
            # Multiple Purchase orders section
            st.write("📋 **Órdenes de compra** *")
            
            # Display current orden de compra inputs
            orden_compra_values = []
            for i, orden in enumerate(st.session_state.orden_compra_list):
                if len(st.session_state.orden_compra_list) == 1:
                    # Single order - full width
                    orden_value = st.text_input(
                        f"Orden {i+1}",
                        value=orden,
                        placeholder=f"Ej: OC-2024-00{i+1}",
                        key=f"orden_{i}"
                    )
                    orden_compra_values.append(orden_value)
                else:
                    # Multiple orders - use columns for remove button
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        orden_value = st.text_input(
                            f"Orden {i+1}",
                            value=orden,
                            placeholder=f"Ej: OC-2024-00{i+1}",
                            key=f"orden_{i}"
                        )
                        orden_compra_values.append(orden_value)
                    with col2:
                        st.write("")  # Empty space for alignment
                        if st.button("🗑️", key=f"remove_{i}"):
                            st.session_state.orden_compra_list.pop(i)
                            st.rerun()
            
            # Update session state with current values
            st.session_state.orden_compra_list = orden_compra_values
            
            # Add button
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("➕ Agregar otra orden", use_container_width=True):
                    st.session_state.orden_compra_list.append('')
                    st.rerun()
            
            # Confirm button
            if st.button("✅ Confirmar Reserva", use_container_width=True):
                # Filter out empty orders and validate
                valid_orders = [orden.strip() for orden in orden_compra_values if orden.strip()]
                
                if valid_orders:
                    # Join multiple orders with comma
                    orden_compra_combined = ', '.join(valid_orders)
                    
                    # UPDATED: Modified format for Fecha and Hora
                    new_booking = {
                        'Fecha': selected_date.strftime('%Y-%m-%d') + ' 00:00:00',
                        'Hora': st.session_state.selected_slot + ':00',
                        'Proveedor': st.session_state.supplier_name,
                        'Numero_de_bultos': numero_bultos,
                        'Orden_de_compra': orden_compra_combined
                    }
                    
                    with st.spinner("Guardando reserva..."):
                        success = save_booking_to_excel(new_booking)
                    
                    if success:
                        st.success("✅ Reserva confirmada!")
                        
                        # Force refresh data to verify save
                        with st.spinner("Verificando reserva..."):
                            download_excel_to_memory.clear()
                            _, updated_reservas_df, _ = download_excel_to_memory()
                            
                        # Verify the booking was saved
                        if updated_reservas_df is not None:
                            saved_booking = updated_reservas_df[
                                (updated_reservas_df['Fecha'] == new_booking['Fecha']) & 
                                (updated_reservas_df['Hora'] == new_booking['Hora']) & 
                                (updated_reservas_df['Proveedor'] == new_booking['Proveedor'])
                            ]
                            if not saved_booking.empty:
                                st.info("✅ Reserva verificada en Excel")
                            else:
                                st.error("❌ Reserva no encontrada en Excel después de guardar")
                        
                        # Send email if email is available
                        if st.session_state.supplier_email:
                            with st.spinner("Enviando confirmación por email..."):
                                email_sent, actual_cc_emails = send_booking_email(
                                    st.session_state.supplier_email,
                                    st.session_state.supplier_name,
                                    new_booking,
                                    st.session_state.supplier_cc_emails  # Pass CC emails from session
                                )
                            if email_sent:
                                st.success(f"📧 Email de confirmación enviado a: {st.session_state.supplier_email}")
                                if actual_cc_emails:
                                    st.success(f"📧 CC enviado a: {', '.join(actual_cc_emails)}")
                            else:
                                st.warning("⚠️ Reserva guardada pero error enviando email")
                        else:
                            st.warning("⚠️ No se encontró email para enviar confirmación")
                        
                        st.balloons()
                        
                        # Clear orden de compra list and log off user
                        st.session_state.orden_compra_list = ['']
                        st.info("Cerrando sesión automáticamente...")
                        st.session_state.authenticated = False
                        st.session_state.supplier_name = None
                        st.session_state.supplier_email = None
                        st.session_state.supplier_cc_emails = []
                        if 'selected_slot' in st.session_state:
                            del st.session_state.selected_slot
                        
                        # Wait a moment then rerun
                        import time
                        time.sleep(2)
                        st.rerun()
                    else:
                        st.error("❌ Error al guardar reserva")
                else:
                    st.error("❌ Al menos una orden de compra es obligatoria")

if __name__ == "__main__":
    main()