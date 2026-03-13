import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Protection",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Sidebar navigation
st.sidebar.title("🛡️ Protection")
st.sidebar.markdown("---")
page = st.sidebar.radio("Navigate", ["Home", "Dashboard", "Settings"])

# --- Pages ---
if page == "Home":
    st.title("Welcome to Protection")
    st.markdown(
        """
        This is the **Protection** application.

        Use the sidebar to navigate between sections.
        """
    )

elif page == "Dashboard":
    st.title("Dashboard")
    st.info("Dashboard content goes here.")

elif page == "Settings":
    st.title("Settings")
    st.info("Settings content goes here.")
