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
@st.cache_data(ttl=300)  # Cache for 5 minutes
def download_excel_to_memory():
    """Download Excel file from SharePoint to memory - INCLUDES ALL SHEETS"""
    try:
        # Authenticate
        user_credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SITE_URL).with_credentials(user_credentials)
        
        # Get file
        file = ctx.web.get_file_by_id(FILE_ID)
        ctx.load(file)
        ctx.execute_query()
        
        # Download to memory
        file_content = io.BytesIO()
        
        # Try multiple download methods
        try:
            file.download(file_content)
            ctx.execute_query()
        except TypeError:
            try:
                response = file.download()
                ctx.execute_query()
                file_content = io.BytesIO(response.content)
            except:
                file.download_session(file_content)
                ctx.execute_query()
        
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
        return None, None, None

def save_booking_to_excel(new_booking):
    """Save new booking to Excel file - PRESERVES ALL SHEETS"""
    try:
        # Load current data - UPDATED TO LOAD ALL SHEETS
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()
        
        if reservas_df is None:
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
        
        # Clear cache
        download_excel_to_memory.clear()
        
        return True
        
    except Exception as e:
        st.error(f"Error guardando reserva: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 3. Email Functions
# ─────────────────────────────────────────────────────────────
def send_booking_email(supplier_email, supplier_name, booking_details):
    """Send booking confirmation email"""
    try:
        # Default CC recipients
        cc_emails = ["leonardo.byon@gmail.com"]
        
        # Email content
        subject = "Confirmación de Reserva para Entrega de Mercadería"
        
        body = f"""
        Hola {supplier_name},
        
        Su reserva de entrega ha sido confirmada exitosamente.
        
        DETALLES DE LA RESERVA:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        📅 Fecha: {booking_details['Fecha']}
        🕐 Horario: {booking_details['Hora']}
        📦 Número de bultos: {booking_details['Numero_de_bultos']}
        📋 Orden de compra: {booking_details['Orden_de_compra']}
        
        INSTRUCCIONES:
        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        • Llegue puntualmente en el horario reservado
        • Tenga lista el Orden de Compra y cualquier otra documentación relevante
        • Asegúrese de que los productos y numero de bultos coincidan con el Orden de Compra
        • Si llega tarde, posiblemente tendra que esperar hasta el proximo cupo disponible del dia
        
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
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Send email
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        
        # Send to supplier + CC recipients
        all_recipients = [supplier_email] + cc_emails
        text = msg.as_string()
        server.sendmail(EMAIL_USER, all_recipients, text)
        server.quit()
        
        return True
        
    except Exception as e:
        st.error(f"Error enviando email: {str(e)}")
        return False

# ─────────────────────────────────────────────────────────────
# 4. Time Slot Functions
# ─────────────────────────────────────────────────────────────
def generate_time_slots():
    """Generate available time slots"""
    # Monday-Friday: 9:00-16:00, Saturday: 9:00-12:00
    weekday_slots = []
    saturday_slots = []
    
    # Weekday slots (9:00-16:00)
    start_hour = 9
    end_hour = 16
    for hour in range(start_hour, end_hour):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            end_minute = minute + 30
            end_hour_calc = hour if end_minute < 60 else hour + 1
            end_minute = end_minute if end_minute < 60 else 0
            end_time = f"{end_hour_calc:02d}:{end_minute:02d}"
            weekday_slots.append(f"{start_time}-{end_time}")
    
    # Saturday slots (9:00-12:00)
    for hour in range(9, 12):
        for minute in [0, 30]:
            start_time = f"{hour:02d}:{minute:02d}"
            end_minute = minute + 30
            end_hour_calc = hour if end_minute < 60 else hour + 1
            end_minute = end_minute if end_minute < 60 else 0
            end_time = f"{end_hour_calc:02d}:{end_minute:02d}"
            saturday_slots.append(f"{start_time}-{end_time}")
    
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
    """Authenticate user against Excel data and get email"""
    credentials_df, _, _ = download_excel_to_memory()  # UPDATED - Now returns 3 values
    
    if credentials_df is None:
        return False, "Error al cargar credenciales", None
    
    # Clean and compare (all data is already strings)
    df_usuarios = credentials_df['usuario'].str.strip()
    
    input_usuario = str(usuario).strip()
    input_password = str(password).strip()
    
    # Find user row
    user_row = credentials_df[df_usuarios == input_usuario]
    if user_row.empty:
        return False, "Usuario no encontrado", None
    
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
        
        return True, "Autenticación exitosa", email
    
    return False, "Contraseña incorrecta", None

# ─────────────────────────────────────────────────────────────
# 6. Main App - UPDATED TO USE ALL SHEETS
# ─────────────────────────────────────────────────────────────
def main():
    st.title("🚚 Dismac: Reserva de Entrega de Mercadería")
    
    # Download Excel when app starts - UPDATED
    with st.spinner("Cargando datos..."):
        credentials_df, reservas_df, gestion_df = download_excel_to_memory()  # UPDATED - Now gets 3 values
    
    if credentials_df is None:
        st.error("❌ Error al cargar archivo")
        return
    
    #st.success(f"✅ Datos cargados: {len(credentials_df)} usuarios, {len(reservas_df)} reservas")
    
    # Session state
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'supplier_name' not in st.session_state:
        st.session_state.supplier_name = None
    if 'supplier_email' not in st.session_state:
        st.session_state.supplier_email = None
    
    # Authentication
    if not st.session_state.authenticated:
        st.subheader("🔐 Iniciar Sesión")
        
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Iniciar Sesión")
            
            if submitted:
                if usuario and password:
                    is_valid, message, email = authenticate_user(usuario, password)
                    
                    if is_valid:
                        st.session_state.authenticated = True
                        st.session_state.supplier_name = usuario
                        st.session_state.supplier_email = email
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
                st.rerun()
        
        st.markdown("---")
        
        # Date selection
        st.subheader("📅 Seleccionar Fecha")
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
        
        # Generate all slots and check availability
        weekday_slots, saturday_slots = generate_time_slots()
        
        if selected_date.weekday() == 5:  # Saturday
            all_slots = saturday_slots
        else:  # Monday-Friday
            all_slots = weekday_slots
        
        # Get booked slots for this date
        date_str = selected_date.strftime('%Y-%m-%d')
        booked_slots = reservas_df[reservas_df['Fecha'] == date_str]['Hora'].tolist()
        
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
        
        # Booking form
        if selected_slot or 'selected_slot' in st.session_state:
            if selected_slot:
                st.session_state.selected_slot = selected_slot
            
            st.markdown("---")
            st.subheader("📦 Información de Entrega")
            
            with st.form("booking_form"):
                # Date and time info (full width)
                st.info(f"📅 Fecha: {selected_date}")
                st.info(f"🕐 Horario: {st.session_state.selected_slot}")
                
                # Number of bultos (full width)
                numero_bultos = st.number_input(
                    "📦 Número de bultos", 
                    min_value=1, 
                    value=1,
                    help="Cantidad de bultos o paquetes a entregar"
                )
                
                # Purchase order (full width, mandatory)
                orden_compra = st.text_input(
                    "📋 Orden de compra *", 
                    placeholder="Ej: OC-2024-001",
                    help="Campo obligatorio - Ingrese el número de orden de compra"
                )
                
                submitted = st.form_submit_button("✅ Confirmar Reserva", use_container_width=True)
                
                if submitted:
                    if orden_compra.strip():
                        new_booking = {
                            'Fecha': selected_date.strftime('%Y-%m-%d'),
                            'Hora': st.session_state.selected_slot,
                            'Proveedor': st.session_state.supplier_name,
                            'Numero_de_bultos': numero_bultos,
                            'Orden_de_compra': orden_compra.strip()
                        }
                        
                        with st.spinner("Guardando reserva..."):
                            success = save_booking_to_excel(new_booking)
                        
                        if success:
                            st.success("✅ Reserva confirmada!")
                            
                            # Send email if email is available
                            if st.session_state.supplier_email:
                                with st.spinner("Enviando confirmación por email..."):
                                    email_sent = send_booking_email(
                                        st.session_state.supplier_email,
                                        st.session_state.supplier_name,
                                        new_booking
                                    )
                                if email_sent:
                                    st.success(f"📧 Email de confirmación enviado a: {st.session_state.supplier_email}")
                                else:
                                    st.warning("⚠️ Reserva guardada pero error enviando email")
                            else:
                                st.warning("⚠️ No se encontró email para enviar confirmación")
                            
                            st.balloons()
                            
                            # Log off user and clear session
                            st.info("Cerrando sesión automáticamente...")
                            st.session_state.authenticated = False
                            st.session_state.supplier_name = None
                            st.session_state.supplier_email = None
                            if 'selected_slot' in st.session_state:
                                del st.session_state.selected_slot
                            
                            # Wait a moment then rerun
                            import time
                            time.sleep(2)
                            st.rerun()
                        else:
                            st.error("❌ Error al guardar reserva")
                    else:
                        st.warning("⚠️ Ingrese la orden de compra")

if __name__ == "__main__":
    main()