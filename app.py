import streamlit as st
import sympy as sp
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import schemdraw
import schemdraw.elements as elm
import io
import os

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# 1. Konfigurasi Halaman & Tema Light Mode
st.set_page_config(layout="wide", page_title="Analisis Transien Rangkaian Listrik", page_icon="⚡")

# Fungsi-fungsi SchemDraw (Menyimpan sebagai PNG untuk Support PPTX & Light Theme)
def draw_full_schematic():
    path = "schema_full.svg"
    with schemdraw.Drawing(show=False) as d:
        d.config(color='black', bgcolor='white', lw=2)
        d += elm.SourceSin().up().label('$v(t)$')
        d += elm.Inductor().right().label('$L_1=2H$')
        d += elm.Resistor().right().label('$R_1=4\\Omega$')
        d.push()
        d += elm.Resistor().down().label('$R_{shared}=8\\Omega$', loc='bottom')
        d += elm.Line().left().tox(d.elements[0].start)
        d.pop()
        d += elm.Inductor().right().label('$L_2=4H$')
        d += elm.Resistor().down().label('$R_2=8\\Omega$', loc='bottom')
        d += elm.Line().left().tox(d.elements[3].start)
        try:
            d.save(path)
        except Exception:
            pass
    return path

def draw_loop1_only():
    path = "schema_loop1.svg"
    with schemdraw.Drawing(show=False) as d:
        d.config(color='black', bgcolor='white', lw=2)
        d += elm.SourceSin().up().label('$v(t)$')
        d += elm.Inductor().right().label('$L_1=2H$')
        d += elm.Resistor().right().label('$R_1=4\\Omega$')
        d.push()
        d += elm.Resistor().down().label('$R_{shared}=8\\Omega$', loc='bottom')
        d += elm.Line().left().tox(d.elements[0].start)
        d.pop()
        # Loop 2 Terbuka
        d += elm.Line().right().length(1)
        d += elm.Dot(open=True)
        try:
            d.save(path)
        except Exception:
            pass
    return path

def draw_loop2_only():
    path = "schema_loop2.svg"
    with schemdraw.Drawing(show=False) as d:
        d.config(color='black', bgcolor='white', lw=2)
        d += elm.Dot(open=True)
        d += elm.Line().right().length(1)
        d.push()
        d += elm.Resistor().down().label('$R_{shared}=8\\Omega$', loc='bottom')
        d += elm.Line().left().tox(d.elements[0].start)
        d.pop()
        d += elm.Inductor().right().label('$L_2=4H$')
        d += elm.Resistor().down().label('$R_2=8\\Omega$', loc='bottom')
        d += elm.Line().left().tox(d.elements[3].start)
        try:
            d.save(path)
        except Exception:
            pass
    return path

def create_presentation(fig_arus):
    prs = Presentation()
    
    # Slide 1: Judul
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Analisis Rangkaian Transien Dua Loop"
    slide.placeholders[1].text = "Parameter:\nv(t) = 390cos(t)\ni(0) = 0, i'(0) = 0\nLoop 1: L1=2H, R1=4\u03A9, Rsh=8\u03A9\nLoop 2: L2=4H, R2=8\u03A9"
    
    # Slide 2: Skematik
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Skematik Rangkaian (Loop 1 & Loop 2)"
    try:
        if os.path.exists("schema_loop1.png"):
            slide.shapes.add_picture("schema_loop1.png", Inches(0.5), Inches(2.5), width=Inches(4))
        if os.path.exists("schema_loop2.png"):
            slide.shapes.add_picture("schema_loop2.png", Inches(5), Inches(2.5), width=Inches(4))
    except Exception as e:
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
        txBox.text_frame.text = f"Gambar skematik gagal dimuat: {e}"
        
    # Slide 3: Kesimpulan Soal 1
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Penyelesaian Soal 1 (Sumber Kontinu)"
    tf = slide.placeholders[1].text_frame
    tf.text = "Fungsi Arus dalam Domain Waktu:"
    p = tf.add_paragraph()
    p.text = "i1(t) = -26 \u00B7 exp(-2t) - 16 \u00B7 exp(-8t) + 42 \u00B7 cos(t) + 15 \u00B7 sin(t)"
    p = tf.add_paragraph()
    p.text = "i2(t) = -26 \u00B7 exp(-2t) + 8 \u00B7 exp(-8t) + 18 \u00B7 cos(t) + 12 \u00B7 sin(t)"
    
    # Slide 4: Kesimpulan Soal 2
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Penyelesaian Soal 2 (Sumber Terputus di t=\u03C0)"
    tf = slide.placeholders[1].text_frame
    tf.text = "Dengan Time-Shifting Laplace Transform, forced response menghilang di t \u2265 \u03C0:"
    p = tf.add_paragraph()
    p.text = "i1_putus(t) = i1(t) + i1(t-\u03C0)u(t-\u03C0)"
    p = tf.add_paragraph()
    p.text = "i2_putus(t) = i2(t) + i2(t-\u03C0)u(t-\u03C0)"
    
    # Slide 5: Grafik Respons
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Dinamika Arus (Soal 1 & Soal 2)"
    try:
        img_bytes = fig_arus.to_image(format="png", width=900, height=450)
        img_stream = io.BytesIO(img_bytes)
        slide.shapes.add_picture(img_stream, Inches(0.5), Inches(2), width=Inches(9))
    except Exception as e:
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        txBox.text_frame.text = f"Gagal mengekspor grafik Plotly.\nPastikan library 'kaleido' terinstall: pip install -U kaleido\nError detail: {e}"
        
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- UI Setup ---
st.title("⚡ Laporan Teknis Digital: Analisis Rangkaian Listrik Transien")
st.markdown("---")

def render_latex(formula):
    st.latex(r"\displaystyle " + formula)


# --- 1. Eksposisi Masalah ---
st.header("1. Eksposisi Masalah")
col_exp1, col_exp2 = st.columns([1.5, 1])
with col_exp1:
    st.markdown("""
    Kita akan menganalisis sebuah rangkaian dua loop dengan parameter-parameter berikut:
    - **Tegangan Sumber**: $v(t) = 390 \\cos(t)$ V
    - **Kondisi Awal**: $i(0) = 0$, $i'(0) = 0$
    - **Loop 1**: $R_1 = 4\\,\\Omega$, $L_1 = 2\\text{ H}$, dan resistor bersama $R_{shared} = 8\\,\\Omega$
    - **Loop 2**: $R_2 = 8\\,\\Omega$, $L_2 = 4\\text{ H}$, dan resistor bersama $R_{shared} = 8\\,\\Omega$

    Analisis ini dilakukan menggunakan **Transformasi Laplace** untuk mengekstraksi respons arus secara menyeluruh, mencakup *natural response* maupun *forced response*.
    """)
with col_exp2:
    try:
        path_full = draw_full_schematic()
        with open(path_full, "r") as f:
            st.markdown(f'<div style="text-align:center; padding:10px; background:white; border-radius:10px;">{f.read()}</div>', unsafe_allow_html=True)
    except:
        st.warning("Skematik utama gagal dimuat.")
st.markdown("---")

# --- 2. Penurunan Matematis (Deep Dive) ---
st.header("2. Penurunan Matematis Langkah-Demi-Langkah")

# Setup Symbolic Variables SymPy
t, s = sp.symbols('t s')
I1, I2 = sp.symbols('I_1 I_2')

st.markdown("Pendekatan pedagogis kita memecah rangkaian menjadi dua loop independen untuk mempermudah analisis KVL awal:")

col_L1, col_L2 = st.columns(2)
with col_L1:
    st.markdown("**Analisis Loop 1 (Kiri)**")
    try:
        path_l1 = draw_loop1_only()
        with open(path_l1, "r") as f:
            st.markdown(f'<div style="text-align:center; padding:10px; background:white; border-radius:10px;">{f.read()}</div>', unsafe_allow_html=True)
    except:
        st.warning("Gagal memuat skema Loop 1.")
    render_latex(r"L_1 \frac{di_1}{dt} + R_1 i_1 + R_{sh}(i_1 - i_2) = v(t)")
    render_latex(r"2 \frac{di_1}{dt} + 12i_1 - 8i_2 = v(t)")

with col_L2:
    st.markdown("**Analisis Loop 2 (Kanan)**")
    try:
        path_l2 = draw_loop2_only()
        with open(path_l2, "r") as f:
            st.markdown(f'<div style="text-align:center; padding:10px; background:white; border-radius:10px;">{f.read()}</div>', unsafe_allow_html=True)
    except:
        st.warning("Gagal memuat skema Loop 2.")
    render_latex(r"L_2 \frac{di_2}{dt} + R_2 i_2 + R_{sh}(i_2 - i_1) = 0")
    render_latex(r"4 \frac{di_2}{dt} + 16i_2 - 8i_1 = 0")

st.markdown("---")
# ==========================================
# SOAL 1
# ==========================================
st.subheader("Penyelesaian Soal 1: Sumber Tegangan Kontinu")
st.markdown(r"Pada soal pertama, asumsi tegangan eksitasi tidak pernah terputus, yaitu $v(t) = 390\cos(t)$. Transformasi Laplace untuk sistem ini menghasilkan:")

v_t1 = 390 * sp.cos(t)
V_s1 = sp.laplace_transform(v_t1, t, s, noconds=True).simplify()

render_latex(r"\mathcal{L}\{v(t)\} = " + sp.latex(V_s1))
render_latex(r"\begin{bmatrix} 2s + 12 & -8 \\ -8 & 4s + 16 \end{bmatrix} \begin{bmatrix} I_1(s) \\ I_2(s) \end{bmatrix} = \begin{bmatrix} " + sp.latex(V_s1) + r" \\ 0 \end{bmatrix}")

mat_A = sp.Matrix([[2*s + 12, -8], [-8, 4*s + 16]])
mat_B = sp.Matrix([[V_s1], [0]])
sol = mat_A.LUsolve(mat_B)

I1_s = sp.factor(sol[0])
I2_s = sp.factor(sol[1])

st.markdown("Isolasi variabel arus pada domain-s (Laplace):")
render_latex(r"I_1(s) = " + sp.latex(I1_s))
render_latex(r"I_2(s) = " + sp.latex(I2_s))

st.markdown("Ekspansi pecahan parsial (*Partial Fraction Expansion*):")
I1_pf = sp.apart(I1_s, s).simplify()
I2_pf = sp.apart(I2_s, s).simplify()
render_latex(r"I_1(s) = " + sp.latex(I1_pf))
render_latex(r"I_2(s) = " + sp.latex(I2_pf))

st.markdown("Invers Transformasi Laplace menghasilkan arus $i_1(t)$ dan $i_2(t)$ di Soal 1:")
i1_t = (-26*sp.exp(-2*t) - 16*sp.exp(-8*t) + 42*sp.cos(t) + 15*sp.sin(t)).simplify()
i2_t = (-26*sp.exp(-2*t) + 8*sp.exp(-8*t) + 18*sp.cos(t) + 12*sp.sin(t)).simplify()
render_latex(r"i_{1, Soal1}(t) = " + sp.latex(i1_t))
render_latex(r"i_{2, Soal1}(t) = " + sp.latex(i2_t))

# ==========================================
# SOAL 2
# ==========================================
st.markdown("---")
st.subheader(r"Penyelesaian Soal 2: Sumber Tegangan Terputus di $t=\pi$")
st.markdown(r"Pada kasus ini, tegangan sumber dimatikan secara mendadak pada saat $t = \pi$. Secara matematis, eksitasi dimodelkan dengan fungsi *Heaviside Step Function* $u(t)$:")
render_latex(r"v_2(t) = 390 \cos(t) [1 - u(t - \pi)]")
render_latex(r"v_2(t) = 390 \cos(t) - 390 \cos(t) u(t - \pi)")

st.markdown(r"Mengingat properti fungsi trigonometri $\cos(t - \pi) = -\cos(t)$, kita ubah persamaannya agar selaras dengan **Teorema Pergeseran Waktu (Time-Shifting)**:")
render_latex(r"v_2(t) = 390 \cos(t) + 390 \cos(t - \pi) u(t - \pi)")

st.markdown("Menerapkan Transformasi Laplace menghasilkan komponen *exponential decay* di s-plane:")
render_latex(r"V_2(s) = \frac{390s}{s^2 + 1} + e^{-\pi s} \frac{390s}{s^2 + 1} = \frac{390s}{s^2 + 1} (1 + e^{-\pi s})")

st.markdown("Sehingga fungsi arus pada domain-s (Laplace) untuk Soal 2 secara identik memuat pergeseran yang sama terhadap matriks Soal 1:")
render_latex(r"I_{1, Soal2}(s) = I_1(s) + I_1(s) e^{-\pi s}")
render_latex(r"I_{2, Soal2}(s) = I_2(s) + I_2(s) e^{-\pi s}")

st.markdown("Dengan menerapkan *Inverse Laplace* pada teorema tersebut, hasil di domain waktu merepresentasikan redaman asimtotik yang murni:")
render_latex(r"i_{1, Soal2}(t) = i_1(t) + i_1(t - \pi) u(t - \pi)")
render_latex(r"i_{2, Soal2}(t) = i_2(t) + i_2(t - \pi) u(t - \pi)")
st.markdown(r"*(Penting: Saat $t \geq \pi$, seluruh komponen sinusoidal atau forced response akan saling membatalkan secara matematis, menyisakan respons energi buangan (*natural decay*) pada resistor.)*")

st.markdown("---")

# --- 3. Analisis Stabilitas: Pole-Zero Plot ---
st.header("3. Analisis Stabilitas: Pole-Zero Plot")
col_pz1, col_pz2 = st.columns([1.5, 1])

with col_pz1:
    fig_pz = go.Figure()
    # Poles: dari penyebut (s+2)(s+8)(s^2+1) -> -2, -8, +j, -j
    poles_real = [-2, -8, 0, 0]
    poles_imag = [0, 0, 1, -1]
    # Zeros for I1(s) -> s=0, s=-4
    zeros_real = [0, -4]
    zeros_imag = [0, 0]
    
    fig_pz.add_trace(go.Scatter(x=poles_real, y=poles_imag, mode='markers', 
                                marker=dict(symbol='x', color='red', size=12, line_width=2), 
                                name='Poles'))
    fig_pz.add_trace(go.Scatter(x=zeros_real, y=zeros_imag, mode='markers', 
                                marker=dict(symbol='circle-open', color='blue', size=10, line_width=2), 
                                name='Zeros (I1)'))
    
    fig_pz.add_vline(x=0, line_dash="solid", line_color="black", line_width=1.5)
    fig_pz.add_hline(y=0, line_dash="solid", line_color="black", line_width=1.5)
    
    fig_pz.update_layout(
        title="Pole-Zero Map (s-plane)",
        xaxis_title="Real (\u03C3)",
        yaxis_title="Imaginary (j\u03C9)",
        template="plotly_white",
        width=600, height=400,
        xaxis=dict(range=[-10, 2]),
        yaxis=dict(range=[-2, 2])
    )
    st.plotly_chart(fig_pz, width="stretch")

with col_pz2:
    st.markdown(r"""
    **Analisis Stabilitas Sistem:**
    - **Poles di Sumbu Imajiner ($\pm j$ atau $+j$ dan $-j$)**: Berasal dari fungsi eksitasi input tegangan. Poles ini merepresentasikan respons osilasi *steady-state* yang konstan.
    - **Poles di Sumbu Real Negatif (-2 dan -8)**: Menunjukkan karakteristik intrinsik rangkaian (R dan L). Letaknya berada di *Left-Half Plane* (LHP), yang mengartikan bahwa sistem bersifat **stabil secara asimtotik**. Tanpa paksaan dari luar, respons alamiahnya akan berupa peluruhan eksponensial ($e^{-2t}$ dan $e^{-8t}$) menuju angka nol murni.
    """)

st.markdown("---")

# --- 4. Analisis Respons Transien & Visualisasi ---
st.header("4. Analisis Respons Transien")

# Perpanjang waktu hingga 25 detik
t_vals = np.linspace(0, 25, 1000)

# Array Soal 1: Kontinu
i1_cont = -26*np.exp(-2*t_vals) - 16*np.exp(-8*t_vals) + 42*np.cos(t_vals) + 15*np.sin(t_vals)
i2_cont = -26*np.exp(-2*t_vals) + 8*np.exp(-8*t_vals) + 18*np.cos(t_vals) + 12*np.sin(t_vals)

# Array Soal 2: Putus pada t=pi
i1_putus = np.where(t_vals < np.pi, 
                    i1_cont, 
                    -26*np.exp(-2*t_vals)*(1 + np.exp(2*np.pi)) - 16*np.exp(-8*t_vals)*(1 + np.exp(8*np.pi)))
i2_putus = np.where(t_vals < np.pi,
                    i2_cont,
                    -26*np.exp(-2*t_vals)*(1 + np.exp(2*np.pi)) + 8*np.exp(-8*t_vals)*(1 + np.exp(8*np.pi)))

fig_arus = make_subplots(rows=1, cols=2, subplot_titles=("Arus Loop 1 (i1)", "Arus Loop 2 (i2)"), shared_yaxes=True)

# Plot Loop 1
fig_arus.add_trace(go.Scatter(x=t_vals, y=i1_cont, mode='lines', name='Soal 1 (i1 Kontinu)', line=dict(color='#1f77b4', width=2)), row=1, col=1)
fig_arus.add_trace(go.Scatter(x=t_vals, y=i1_putus, mode='lines', name='Soal 2 (i1 Putus t=π)', line=dict(color='#ff7f0e', dash='dash', width=2.5)), row=1, col=1)

# Plot Loop 2
fig_arus.add_trace(go.Scatter(x=t_vals, y=i2_cont, mode='lines', name='Soal 1 (i2 Kontinu)', line=dict(color='#2ca02c', width=2)), row=1, col=2)
fig_arus.add_trace(go.Scatter(x=t_vals, y=i2_putus, mode='lines', name='Soal 2 (i2 Putus t=π)', line=dict(color='#d62728', dash='dash', width=2.5)), row=1, col=2)

# Garis putus pada t = pi
pi_val = np.pi
fig_arus.add_vline(x=pi_val, line_dash="dot", line_color="black", row=1, col=1)
fig_arus.add_vline(x=pi_val, line_dash="dot", line_color="black", row=1, col=2)

fig_arus.update_layout(
    template="plotly_white", 
    title_text="Dinamika Respons Arus Rangkaian (Soal 1 & Soal 2 Side-by-Side)",
    hovermode="x unified",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)
fig_arus.update_xaxes(title_text="Waktu (detik)", row=1, col=1)
fig_arus.update_xaxes(title_text="Waktu (detik)", row=1, col=2)
fig_arus.update_yaxes(title_text="Arus (Ampere)", row=1, col=1)

st.plotly_chart(fig_arus, width="stretch")

st.markdown(r"""
### Narasi Analisis Fisis
**Observasi Grafik:**
1. **Soal 1 (Garis Solid)**: Pada detik awal ($0 \leq t < 3$ detik), terjadi efek *transien* di mana amplitudo belum stabil dan gelombang terdistorsi oleh redaman eksponensial. Setelah melampaui rentang $t > 5$ detik, arus sepenuhnya masuk ke dalam wujud *steady-state* osilasi murni tanpa redaman akibat suplai paksa berkelanjutan dari tegangan sumber.
2. **Soal 2 (Garis Putus-Putus)**: Tepat pada $t = \pi \approx 3.14$ detik (ditandai garis vertikal), interupsi eksitasi terjadi. Secara visual, arus pada grafik tidak jatuh lurus menghantam 0 Ampere, membuktikan inersia akibat induktor $L$. Secara empiris, *forced response* lenyap, menyisakan *natural response* peluruhan eksponensial murni yang meluruh ke arah ekuilibrium. Saat waktu diekspansi hingga $25$ detik, energi magnetik di dalam kumparan dipastikan telah terkuras seutuhnya (menjadi panas oleh resistor $R$).
""")

# --- Export PPTX Section ---
st.markdown("---")
st.subheader("Ekspor ke Presentasi")
if HAS_PPTX:
    try:
        ppt_file = create_presentation(fig_arus)
        st.download_button(
            label="Download Laporan PPTX 📥",
            data=ppt_file,
            file_name="Laporan_Analisis_Transien.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        st.error(f"Gagal memproses PPTX: {e}")
else:
    st.warning("Library `python-pptx` belum terpasang.")
    st.info("Untuk mengaktifkan fitur Export ke PPTX, jalankan perintah di terminal:\n\n`pip install python-pptx kaleido`")
