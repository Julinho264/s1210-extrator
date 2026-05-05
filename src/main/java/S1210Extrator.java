// POI — apenas o necessário, sem Color nem Font (conflito com java.awt)
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// XML
import org.w3c.dom.*;
import javax.xml.parsers.*;

// Swing / AWT
import javax.swing.*;
import javax.swing.border.*;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.BorderLayout;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

// IO / NIO / util / ZIP
import java.io.*;
import java.nio.file.*;
import java.util.Properties;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class S1210Extrator extends JFrame {

    // ── Paleta Oficial SIGEP ──────────────────────────────────────────────────
    static final Color NAVY       = new Color(0x1D, 0x3B, 0x60);  // fundo cabeçalho
    static final Color BLUE       = new Color(0x25, 0x71, 0xE8);  // azul primário
    static final Color DEEP_BLUE  = new Color(0x00, 0x49, 0xBA);  // hover / acento
    static final Color LIGHT_BLUE = new Color(0xB1, 0xC6, 0xDD);  // bordas suaves
    static final Color BG         = new Color(0xF2, 0xF4, 0xF7);  // fundo geral
    static final Color CARD       = Color.WHITE;
    static final Color GRAY_TEXT  = new Color(0x70, 0x80, 0x90);

    // ── Fontes Figtree ────────────────────────────────────────────────────────
    private static final java.awt.Font FIGTREE_REGULAR;
    private static final java.awt.Font FIGTREE_SEMI;
    private static final java.awt.Font FIGTREE_BOLD;

    static {
        FIGTREE_REGULAR = loadFigtree("/fonts/Figtree-Regular.ttf");
        FIGTREE_SEMI    = loadFigtree("/fonts/Figtree-SemiBold.ttf");
        FIGTREE_BOLD    = loadFigtree("/fonts/Figtree-Bold.ttf");
    }

    private static java.awt.Font loadFigtree(String res) {
        try (InputStream is = S1210Extrator.class.getResourceAsStream(res)) {
            if (is == null) return new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12);
            return java.awt.Font.createFont(java.awt.Font.TRUETYPE_FONT, is);
        } catch (Exception e) {
            return new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 12);
        }
    }

    // ── Componentes de UI ─────────────────────────────────────────────────────
    private JTextField  txtPasta;
    private JTextField  txtSaida;
    private JButton     btnSelecionarPasta;
    private JButton     btnSelecionarSaida;
    private JButton     btnGerar;
    private JButton     btnCancelar;
    private JTextArea   txtLog;
    private JProgressBar progressBar;
    private JLabel      lblStatus;
    private JCheckBox   chkDivergencias;

    // ── Buffer de log — evita inundar a EDT com invokeLater por arquivo ───────
    private final java.util.concurrent.ConcurrentLinkedQueue<String> logBuffer =
        new java.util.concurrent.ConcurrentLinkedQueue<>();
    private int ultimoPctBar = -1; // controla atualização da barra de progresso

    // ── Cancelamento do processamento ─────────────────────────────────────────
    private volatile boolean cancelado = false;

    // ── Versão (injetada pelo Maven via app.properties) ───────────────────────
    static final String VERSAO = lerVersao();

    private static String lerVersao() {
        try (InputStream is = S1210Extrator.class.getResourceAsStream("/app.properties")) {
            if (is == null) return "?.?";
            Properties props = new Properties();
            props.load(is);
            return props.getProperty("version", "?.?");
        } catch (Exception e) { return "?.?"; }
    }

    // =========================================================================
    public S1210Extrator() {
        super("SIGEP — Extrator S-1210  v" + VERSAO);
        initUI();
    }

    // ── Construção da Interface ───────────────────────────────────────────────
    private void initUI() {
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(760, 560);
        setMinimumSize(new Dimension(620, 480));
        setLocationRelativeTo(null);

        ImageIcon ico = loadLogo(64, 64, false);
        if (ico != null) setIconImage(ico.getImage());

        // Branding dos diálogos JOptionPane
        UIManager.put("OptionPane.background",        CARD);
        UIManager.put("Panel.background",             CARD);
        UIManager.put("OptionPane.messageForeground", NAVY);
        UIManager.put("OptionPane.messageFont",       FIGTREE_REGULAR.deriveFont(13f));
        UIManager.put("OptionPane.buttonFont",        FIGTREE_SEMI.deriveFont(12f));
        UIManager.put("OptionPane.minimumSize",       new Dimension(360, 0));

        JPanel root = new JPanel(new BorderLayout());
        root.setBackground(BG);

        root.add(buildHeader(),  BorderLayout.NORTH);
        root.add(buildContent(), BorderLayout.CENTER);
        root.add(buildFooter(),  BorderLayout.SOUTH);

        setContentPane(root);
    }

    // ── Cabeçalho (fundo escuro → logo negativa) ──────────────────────────────
    private JPanel buildHeader() {
        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.setBackground(NAVY);

        JPanel header = new JPanel(new BorderLayout(0, 0));
        header.setBackground(NAVY);
        header.setBorder(new EmptyBorder(14, 20, 14, 20));

        // Logo negativa (branca) sobre fundo escuro NAVY — desenhada via G2D para respeitar DPI
        JLabel logoLabel = makeLogoLabel(220, 52, true);
        logoLabel.setForeground(Color.WHITE);
        header.add(logoLabel, BorderLayout.WEST);

        // Título da ferramenta à direita
        JPanel titlePanel = new JPanel(new GridLayout(2, 1));
        titlePanel.setBackground(NAVY);

        JLabel lblTool = new JLabel("Extrator S-1210", SwingConstants.RIGHT);
        lblTool.setFont(FIGTREE_BOLD.deriveFont(java.awt.Font.BOLD, 17f));
        lblTool.setForeground(Color.WHITE);

        JLabel lblSub = new JLabel("eSocial — Pagamentos de Rendimentos", SwingConstants.RIGHT);
        lblSub.setFont(FIGTREE_REGULAR.deriveFont(11f));
        lblSub.setForeground(LIGHT_BLUE);

        titlePanel.add(lblTool);
        titlePanel.add(lblSub);
        header.add(titlePanel, BorderLayout.EAST);

        // Linha de acento azul na base
        JPanel accent = new JPanel();
        accent.setBackground(BLUE);
        accent.setPreferredSize(new Dimension(0, 4));

        wrapper.add(header, BorderLayout.CENTER);
        wrapper.add(accent, BorderLayout.SOUTH);
        return wrapper;
    }

    // ── Conteúdo central ──────────────────────────────────────────────────────
    private JPanel buildContent() {
        JPanel content = new JPanel(new BorderLayout(0, 12));
        content.setBackground(BG);
        content.setBorder(new EmptyBorder(16, 16, 8, 16));

        JPanel cards = new JPanel(new GridLayout(2, 1, 0, 10));
        cards.setOpaque(false);

        txtPasta = makeTextField();
        txtSaida = makeTextField();
        btnSelecionarPasta = makeSecondaryButton("Selecionar...");
        btnSelecionarSaida = makeSecondaryButton("Selecionar...");
        btnSelecionarPasta.addActionListener(this::selecionarPasta);
        btnSelecionarSaida.addActionListener(this::selecionarSaida);

        cards.add(buildCard("Pasta com arquivos XML S-1210", txtPasta, btnSelecionarPasta));
        cards.add(buildCard("Arquivo de saída (.xlsx)", txtSaida, btnSelecionarSaida));

        // Checkbox de filtro
        chkDivergencias = new JCheckBox("Listar apenas divergências");
        chkDivergencias.setFont(FIGTREE_SEMI.deriveFont(12f));
        chkDivergencias.setForeground(NAVY);
        chkDivergencias.setOpaque(false);
        chkDivergencias.setToolTipText(
            "Inclui somente registros onde perApur não coincide com nenhum perRef " +
            "(exceto referência anual, ex.: 2025 coincide com 2025-12)");

        JPanel optionsRow = new JPanel(new BorderLayout());
        optionsRow.setOpaque(false);
        optionsRow.setBorder(new EmptyBorder(4, 2, 0, 0));
        optionsRow.add(chkDivergencias, BorderLayout.WEST);

        JPanel northBlock = new JPanel(new BorderLayout(0, 6));
        northBlock.setOpaque(false);
        northBlock.add(cards,      BorderLayout.NORTH);
        northBlock.add(optionsRow, BorderLayout.SOUTH);
        content.add(northBlock, BorderLayout.NORTH);

        content.add(buildLogPanel(), BorderLayout.CENTER);
        return content;
    }

    private JPanel buildCard(String titulo, JTextField campo, JButton botao) {
        JPanel card = new JPanel(new BorderLayout(10, 6));
        card.setBackground(CARD);
        card.setBorder(new CompoundBorder(
            new MatteBorder(0, 4, 0, 0, BLUE),
            new CompoundBorder(
                new LineBorder(new Color(0xD8, 0xE4, 0xF0), 1),
                new EmptyBorder(10, 12, 10, 12)
            )
        ));

        JLabel lbl = new JLabel(titulo);
        lbl.setFont(FIGTREE_SEMI.deriveFont(java.awt.Font.BOLD, 12f));
        lbl.setForeground(NAVY);

        card.add(lbl,    BorderLayout.NORTH);
        card.add(campo,  BorderLayout.CENTER);
        card.add(botao,  BorderLayout.EAST);
        return card;
    }

    private JPanel buildLogPanel() {
        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.setBackground(CARD);
        panel.setBorder(new CompoundBorder(
            new MatteBorder(0, 4, 0, 0, LIGHT_BLUE),
            new CompoundBorder(
                new LineBorder(new Color(0xD8, 0xE4, 0xF0), 1),
                new EmptyBorder(10, 12, 10, 12)
            )
        ));

        JLabel lbl = new JLabel("Log de processamento");
        lbl.setFont(FIGTREE_SEMI.deriveFont(java.awt.Font.BOLD, 12f));
        lbl.setForeground(NAVY);

        txtLog = new JTextArea();
        txtLog.setEditable(false);
        txtLog.setFont(new java.awt.Font("Consolas", java.awt.Font.PLAIN, 11));
        txtLog.setForeground(new Color(0x1D, 0x3B, 0x60));
        txtLog.setBackground(new Color(0xF8, 0xFB, 0xFF));
        txtLog.setBorder(new EmptyBorder(4, 6, 4, 6));

        JScrollPane scroll = new JScrollPane(txtLog);
        scroll.setBorder(new LineBorder(new Color(0xD0, 0xDF, 0xEF)));

        panel.add(lbl,    BorderLayout.NORTH);
        panel.add(scroll, BorderLayout.CENTER);
        return panel;
    }

    // ── Rodapé (fundo claro → logo padrão) ───────────────────────────────────
    private JPanel buildFooter() {
        JPanel footer = new JPanel(new BorderLayout(10, 4));
        footer.setBackground(CARD);
        footer.setBorder(new CompoundBorder(
            new MatteBorder(1, 0, 0, 0, LIGHT_BLUE),
            new EmptyBorder(10, 16, 10, 16)
        ));

        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);
        progressBar.setString("Aguardando...");
        progressBar.setForeground(BLUE);
        progressBar.setBackground(new Color(0xE4, 0xEE, 0xF8));
        progressBar.setFont(FIGTREE_REGULAR.deriveFont(11f));
        progressBar.setBorder(new LineBorder(LIGHT_BLUE));
        progressBar.setPreferredSize(new Dimension(0, 26));

        btnGerar = makePrimaryButton("  Gerar Planilha  ");
        btnGerar.addActionListener(e -> iniciarGeracao());

        btnCancelar = makeSecondaryButton("  Cancelar  ");
        btnCancelar.setVisible(false);
        btnCancelar.addActionListener(e -> {
            cancelado = true;
            btnCancelar.setEnabled(false);
            btnCancelar.setText("  Cancelando...  ");
        });

        JPanel btns = new JPanel(new java.awt.FlowLayout(java.awt.FlowLayout.RIGHT, 6, 0));
        btns.setOpaque(false);
        btns.add(btnCancelar);
        btns.add(btnGerar);

        JPanel actionRow = new JPanel(new BorderLayout(10, 0));
        actionRow.setOpaque(false);
        actionRow.add(progressBar, BorderLayout.CENTER);
        actionRow.add(btns,        BorderLayout.EAST);

        lblStatus = new JLabel("© SIGEP — Sistema Integrado de Gestão Pública  |  v" + VERSAO);
        lblStatus.setFont(FIGTREE_REGULAR.deriveFont(10f));
        lblStatus.setForeground(GRAY_TEXT);

        footer.add(actionRow, BorderLayout.CENTER);
        footer.add(lblStatus, BorderLayout.SOUTH);
        return footer;
    }

    // ── Componentes estilizados ───────────────────────────────────────────────
    private JTextField makeTextField() {
        JTextField tf = new JTextField();
        tf.setEditable(false);
        tf.setBackground(new Color(0xF4, 0xF8, 0xFF));
        tf.setForeground(new Color(0x1D, 0x3B, 0x60));
        tf.setFont(FIGTREE_REGULAR.deriveFont(12f));
        tf.setBorder(new CompoundBorder(
            new LineBorder(LIGHT_BLUE),
            new EmptyBorder(4, 8, 4, 8)
        ));
        return tf;
    }

    private JButton makePrimaryButton(String texto) {
        JButton btn = new JButton(texto);
        btn.setBackground(NAVY);
        btn.setForeground(Color.WHITE);
        btn.setFont(FIGTREE_BOLD.deriveFont(java.awt.Font.BOLD, 13f));
        btn.setFocusPainted(false);
        btn.setBorderPainted(false);
        btn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btn.setBorder(new EmptyBorder(8, 18, 8, 18));
        btn.setOpaque(true);
        btn.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent e) { btn.setBackground(DEEP_BLUE); }
            public void mouseExited (MouseEvent e) { btn.setBackground(NAVY); }
        });
        return btn;
    }

    private JButton makeSecondaryButton(String texto) {
        JButton btn = new JButton(texto);
        btn.setBackground(BLUE);
        btn.setForeground(Color.WHITE);
        btn.setFont(FIGTREE_REGULAR.deriveFont(12f));
        btn.setFocusPainted(false);
        btn.setBorderPainted(false);
        btn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        btn.setBorder(new EmptyBorder(6, 14, 6, 14));
        btn.setOpaque(true);
        btn.addMouseListener(new MouseAdapter() {
            public void mouseEntered(MouseEvent e) { btn.setBackground(DEEP_BLUE); }
            public void mouseExited (MouseEvent e) { btn.setBackground(BLUE); }
        });
        return btn;
    }

    // ── Diálogos com identidade SIGEP ─────────────────────────────────────────
    private ImageIcon dialogIcon() { return loadLogo(48, 48, false); }

    private void dlgInfo(String titulo, String msg) {
        JOptionPane.showMessageDialog(this, msg, titulo,
            JOptionPane.INFORMATION_MESSAGE, dialogIcon());
    }
    private void dlgAviso(String titulo, String msg) {
        JOptionPane.showMessageDialog(this, msg, titulo,
            JOptionPane.WARNING_MESSAGE, dialogIcon());
    }
    private void dlgErro(String titulo, String msg) {
        JOptionPane.showMessageDialog(this, msg, titulo,
            JOptionPane.ERROR_MESSAGE, dialogIcon());
    }
    private int dlgConfirmar(String titulo, String msg) {
        return JOptionPane.showConfirmDialog(this, msg, titulo,
            JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, dialogIcon());
    }

    // Desenha a logo diretamente via paintComponent com hints de qualidade,
    // garantindo nitidez em qualquer fator de escala DPI do sistema.
    private JLabel makeLogoLabel(int maxW, int maxH, boolean negative) {
        String res = negative ? "/logo_sigep_negativa.png" : "/logo_sigep.png";
        java.awt.image.BufferedImage loaded;
        try {
            java.net.URL url = getClass().getResource(res);
            loaded = (url != null) ? javax.imageio.ImageIO.read(url) : null;
        } catch (Exception e) { loaded = null; }

        final java.awt.image.BufferedImage img = loaded;

        return new JLabel() {
            @Override
            public Dimension getPreferredSize() {
                if (img == null) return new Dimension(80, maxH);
                double ratio = Math.min((double) maxW / img.getWidth(),
                                        (double) maxH / img.getHeight());
                return new Dimension((int)(img.getWidth() * ratio),
                                     (int)(img.getHeight() * ratio));
            }
            @Override
            protected void paintComponent(java.awt.Graphics g) {
                super.paintComponent(g);
                if (img == null) { g.setColor(java.awt.Color.WHITE); g.drawString("SIGEP", 4, 20); return; }
                java.awt.Graphics2D g2 = (java.awt.Graphics2D) g.create();
                g2.setRenderingHint(java.awt.RenderingHints.KEY_INTERPOLATION,
                                    java.awt.RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                g2.setRenderingHint(java.awt.RenderingHints.KEY_RENDERING,
                                    java.awt.RenderingHints.VALUE_RENDER_QUALITY);
                g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                                    java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
                double ratio = Math.min((double) maxW / img.getWidth(),
                                        (double) maxH / img.getHeight());
                int dw = (int)(img.getWidth() * ratio);
                int dh = (int)(img.getHeight() * ratio);
                g2.drawImage(img, 0, 0, dw, dh, null);
                g2.dispose();
            }
        };
    }

     /**
     * @param negative  true → logo negativa (branca, para fundos escuros);
     *                  false → logo padrão (escura, para fundos claros)
     */
    private ImageIcon loadLogo(int maxW, int maxH, boolean negative) {
        String res = negative ? "/logo_sigep_negativa.png" : "/logo_sigep.png";
        try {
            java.net.URL url = getClass().getResource(res);
            if (url == null) return null;
            java.awt.image.BufferedImage src = javax.imageio.ImageIO.read(url);
            if (src == null) return null;
            int origW = src.getWidth(), origH = src.getHeight();
            double ratio = Math.min((double) maxW / origW, (double) maxH / origH);
            int dstW = (int)(origW * ratio), dstH = (int)(origH * ratio);
            return new ImageIcon(progressiveScale(src, dstW, dstH));
        } catch (Exception e) {
            return null;
        }
    }

    private java.awt.image.BufferedImage progressiveScale(
            java.awt.image.BufferedImage src, int dstW, int dstH) {
        int w = src.getWidth(), h = src.getHeight();
        java.awt.image.BufferedImage cur = src;
        // Halve repeatedly until one more halving would overshoot the target
        while (w > dstW * 2 || h > dstH * 2) {
            w = Math.max(w / 2, dstW);
            h = Math.max(h / 2, dstH);
            cur = scaleStep(cur, w, h);
        }
        return scaleStep(cur, dstW, dstH);
    }

    private java.awt.image.BufferedImage scaleStep(
            java.awt.image.BufferedImage src, int w, int h) {
        java.awt.image.BufferedImage dst =
            new java.awt.image.BufferedImage(w, h,
                java.awt.image.BufferedImage.TYPE_INT_ARGB);
        java.awt.Graphics2D g2 = dst.createGraphics();
        g2.setRenderingHint(java.awt.RenderingHints.KEY_INTERPOLATION,
                            java.awt.RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        g2.setRenderingHint(java.awt.RenderingHints.KEY_RENDERING,
                            java.awt.RenderingHints.VALUE_RENDER_QUALITY);
        g2.setRenderingHint(java.awt.RenderingHints.KEY_ANTIALIASING,
                            java.awt.RenderingHints.VALUE_ANTIALIAS_ON);
        g2.drawImage(src, 0, 0, w, h, null);
        g2.dispose();
        return dst;
    }

    // ── Seleção de caminhos (diálogos nativos do Windows) ────────────────────
    private void selecionarPasta(ActionEvent e) {
        btnSelecionarPasta.setEnabled(false);
        new Thread(() -> {
            String caminho = dialogoPastaNativo();
            SwingUtilities.invokeLater(() -> {
                btnSelecionarPasta.setEnabled(true);
                if (caminho != null) {
                    txtPasta.setText(caminho);
                    if (txtSaida.getText().isBlank())
                        txtSaida.setText(new File(caminho, "s1210_resultado.xlsx").getAbsolutePath());
                }
            });
        }).start();
    }

    private void selecionarSaida(ActionEvent e) {
        btnSelecionarSaida.setEnabled(false);
        new Thread(() -> {
            String caminho = dialogoSalvarNativo();
            SwingUtilities.invokeLater(() -> {
                btnSelecionarSaida.setEnabled(true);
                if (caminho != null) txtSaida.setText(caminho);
            });
        }).start();
    }

    // ── Roteador de plataforma ────────────────────────────────────────────────
    private static final String OS = System.getProperty("os.name", "").toLowerCase();

    private String dialogoPastaNativo() {
        if (OS.contains("win"))  return dialogoPastaWindows();
        if (OS.contains("mac"))  return dialogoComInvoke(this::dialogoPastaMac);
        return dialogoComInvoke(this::dialogoPastaSwing);
    }

    private String dialogoSalvarNativo() {
        if (OS.contains("win"))  return dialogoSalvarWindows();
        if (OS.contains("mac"))  return dialogoComInvoke(this::dialogoSalvarMac);
        return dialogoComInvoke(this::dialogoSalvarSwing);
    }

    /** Executa um diálogo Swing no EDT a partir de uma thread de fundo. */
    private String dialogoComInvoke(java.util.concurrent.Callable<String> fn) {
        String[] result = {null};
        try {
            SwingUtilities.invokeAndWait(() -> {
                try { result[0] = fn.call(); } catch (Exception ignored) {}
            });
        } catch (Exception ignored) {}
        return result[0];
    }

    // ── Windows: PowerShell nativo (Explorer moderno, UTF-8 correto) ─────────
    private String dialogoPastaWindows() {
        try {
            String script =
                "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
                "Add-Type -AssemblyName System.Windows.Forms; " +
                "$d = New-Object System.Windows.Forms.OpenFileDialog; " +
                "$d.Title = 'Selecionar pasta com arquivos XML / ZIP do eSocial'; " +
                "$d.ValidateNames = $false; " +
                "$d.CheckFileExists = $false; " +
                "$d.CheckPathExists = $true; " +
                "$d.FileName = 'Selecionar pasta'; " +
                "if ($d.ShowDialog() -eq 'OK') { Write-Output ([System.IO.Path]::GetDirectoryName($d.FileName)) }";
            Process proc = new ProcessBuilder("powershell.exe", "-NonInteractive", "-Command", script)
                .redirectErrorStream(true).start();
            String out = new String(proc.getInputStream().readAllBytes(),
                java.nio.charset.StandardCharsets.UTF_8).trim();
            proc.waitFor();
            return out.isBlank() ? null : out;
        } catch (Exception ex) { return null; }
    }

    private String dialogoSalvarWindows() {
        try {
            String script =
                "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; " +
                "Add-Type -AssemblyName System.Windows.Forms; " +
                "$d = New-Object System.Windows.Forms.SaveFileDialog; " +
                "$d.Title = 'Salvar planilha como...'; " +
                "$d.Filter = 'Planilha Excel (*.xlsx)|*.xlsx'; " +
                "$d.FileName = 's1210_resultado.xlsx'; " +
                "if ($d.ShowDialog() -eq 'OK') { Write-Output $d.FileName }";
            Process proc = new ProcessBuilder("powershell.exe", "-NonInteractive", "-Command", script)
                .redirectErrorStream(true).start();
            String out = new String(proc.getInputStream().readAllBytes(),
                java.nio.charset.StandardCharsets.UTF_8).trim();
            proc.waitFor();
            return out.isBlank() ? null : out;
        } catch (Exception ex) { return null; }
    }

    // ── macOS: java.awt.FileDialog (integrado ao Finder) ─────────────────────
    private String dialogoPastaMac() {
        System.setProperty("apple.awt.fileDialogForDirectories", "true");
        try {
            java.awt.FileDialog fd = new java.awt.FileDialog(
                (java.awt.Frame) SwingUtilities.getWindowAncestor(this),
                "Selecionar pasta com arquivos XML / ZIP do eSocial",
                java.awt.FileDialog.LOAD);
            fd.setVisible(true);
            String dir  = fd.getDirectory();
            String file = fd.getFile();
            if (dir == null || file == null) return null;
            return new File(dir, file).getAbsolutePath();
        } finally {
            System.setProperty("apple.awt.fileDialogForDirectories", "false");
        }
    }

    private String dialogoSalvarMac() {
        java.awt.FileDialog fd = new java.awt.FileDialog(
            (java.awt.Frame) SwingUtilities.getWindowAncestor(this),
            "Salvar planilha como...",
            java.awt.FileDialog.SAVE);
        fd.setFile("s1210_resultado.xlsx");
        fd.setVisible(true);
        String dir  = fd.getDirectory();
        String file = fd.getFile();
        if (dir == null || file == null) return null;
        String path = new File(dir, file).getAbsolutePath();
        return path.endsWith(".xlsx") ? path : path + ".xlsx";
    }

    // ── Linux / fallback: JFileChooser ───────────────────────────────────────
    private String dialogoPastaSwing() {
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        fc.setDialogTitle("Selecionar pasta com arquivos XML / ZIP do eSocial");
        return fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION
            ? fc.getSelectedFile().getAbsolutePath() : null;
    }

    private String dialogoSalvarSwing() {
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File("s1210_resultado.xlsx"));
        fc.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
            "Planilha Excel (*.xlsx)", "xlsx"));
        fc.setDialogTitle("Salvar planilha como...");
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return null;
        String path = fc.getSelectedFile().getAbsolutePath();
        return path.endsWith(".xlsx") ? path : path + ".xlsx";
    }

    // ── Início do processamento (thread separada) ─────────────────────────────
    private void iniciarGeracao() {
        String pasta = txtPasta.getText().trim();
        String saida = txtSaida.getText().trim();

        if (pasta.isBlank() || saida.isBlank()) {
            dlgAviso("Atenção", "Selecione a pasta de entrada e o arquivo de saída.");
            return;
        }

        cancelado = false;
        ultimoPctBar = -1;
        btnGerar.setEnabled(false);
        btnCancelar.setVisible(true);
        btnCancelar.setEnabled(true);
        btnCancelar.setText("  Cancelar  ");
        txtLog.setText("");
        progressBar.setValue(0);
        progressBar.setString("Aguardando...");

        new Thread(() -> {
            try {
                processar(pasta, saida);
            } catch (Throwable ex) {
                if (!cancelado) {
                    String msg = ex.getClass().getSimpleName() + ": " + ex.getMessage();
                    log("ERRO FATAL: " + msg);
                    flushLog();
                    gravarLogErro(ex);
                    SwingUtilities.invokeLater(() ->
                        dlgErro("Erro", "Erro durante o processamento:\n" + msg +
                            "\n\nDetalhes gravados em s1210_erro.log"));
                }
            } finally {
                SwingUtilities.invokeLater(() -> {
                    btnGerar.setEnabled(true);
                    btnCancelar.setVisible(false);
                    progressBar.setIndeterminate(false);
                    if (cancelado) {
                        progressBar.setValue(0);
                        progressBar.setString("Cancelado");
                    }
                });
            }
        }).start();
    }

    // =========================================================================
    // Estrutura de dados: CPF + perApur → [ideDmDev], [perRef], nrRecibo
    //   perApur         — de <ideEvento><perApur>, único por arquivo
    //   perRef          — de cada <infoPgto><perRef>, sem repetição por beneficiário
    //   nrRecibo        — de <retornoEvento><recibo><nrRecibo>
    //   dhProcessamento — de <retornoEvento><dhProcessamento>; desempata registros
    //                     com mesmo CPF+perApur (prevalece o mais recente)
    // =========================================================================
    private static class Registro {
        String cpf;
        String perApur;
        String nrRecibo;
        String dhProcessamento = ""; // ISO-8601; string-compare é suficiente
        LinkedHashSet<String> idmDevs   = new LinkedHashSet<>();
        LinkedHashSet<String> perRefs   = new LinkedHashSet<>();
        LinkedHashSet<String> origens   = new LinkedHashSet<>(); // arquivos fonte

        void adicionarDmDev(String v) {
            if (v != null && !v.isBlank()) idmDevs.add(v.trim());
        }

        void adicionarPerRef(String v) {
            if (v != null && !v.isBlank()) perRefs.add(v.trim());
        }

        void adicionarOrigem(String v) {
            if (v != null && !v.isBlank()) origens.add(v.trim());
        }

        String idmDevsStr()   { return String.join("; ", idmDevs); }
        String perRefsStr()   { return String.join("; ", perRefs); }
        String origensStr()   { return String.join("; ", origens); }
    }

    // =========================================================================
    // Processamento principal — streaming (sem acumular bytes em memória)
    // =========================================================================
    private void processar(String pastaCaminho, String saidaCaminho) throws Exception {

        // Feedback imediato — barra indeterminada enquanto varre a pasta
        SwingUtilities.invokeLater(() -> {
            progressBar.setIndeterminate(true);
            progressBar.setString("Analisando pasta...");
        });

        // ── Coleta arquivos da pasta ──────────────────────────────────────────
        List<Path> todosArquivos;
        try (var stream = Files.walk(Paths.get(pastaCaminho))) {
            todosArquivos = stream
                .filter(Files::isRegularFile)
                .sorted()
                .collect(Collectors.toList());
        }

        List<Path> xmlDiretos = todosArquivos.stream()
            .filter(p -> p.toString().toLowerCase().endsWith(".xml"))
            .collect(Collectors.toList());
        List<Path> zips = todosArquivos.stream()
            .filter(p -> p.toString().toLowerCase().endsWith(".zip"))
            .collect(Collectors.toList());

        // Nomes dos XMLs diretos — apenas strings para deduplicação (sem bytes)
        Set<String> nomesVistos = new LinkedHashSet<>();
        for (Path xml : xmlDiretos) {
            nomesVistos.add(xml.getFileName().toString().toUpperCase());
        }

        // Total inicial = XMLs diretos; cresce conforme ZIPs são lidos (sem pré-varredura)
        int[] total = { xmlDiretos.size() };
        int[] cnt   = {0};  // arquivos processados
        int[] s1210 = {0};  // S-1210 encontrados
        int   xmlsDeZip = 0;

        log("Encontrados: " + xmlDiretos.size() + " XML(s) + " + zips.size() + " ZIP(s)");
        flushLog();

        // Volta para barra determinada agora que temos a contagem
        SwingUtilities.invokeLater(() -> {
            progressBar.setIndeterminate(false);
            progressBar.setValue(0);
            progressBar.setString("0%");
        });

        Map<String, Registro> mapa = new LinkedHashMap<>();

        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(false);
        // DocumentBuilder reutilizado — criado novo apenas após exceção de parse
        DocumentBuilder[] dbHolder = { dbf.newDocumentBuilder() };
        dbHolder[0].setErrorHandler(null);

        // ── XMLs diretos: lê e processa um por vez ───────────────────────────
        for (Path xml : xmlDiretos) {
            if (cancelado) { log("⛔ Cancelado pelo usuário."); flushLog(); return; }
            cnt[0]++;
            atualizarProgresso(cnt[0], total[0]);
            // Flush e log de progresso a cada 200 arquivos (mantém UI responsiva)
            if (cnt[0] % 200 == 0) {
                log("⏳ " + cnt[0] + "/" + total[0] + " arquivos — S-1210: " + s1210[0]);
                flushLog();
            }
            String nome = xml.getFileName().toString().toUpperCase();
            try (InputStream is = Files.newInputStream(xml)) {
                List<Registro> regs = parsearStream(nome, is, dbHolder[0]);
                if (!regs.isEmpty()) { s1210[0]++; fundirNoMapa(regs, mapa, nome); }
            } catch (Throwable ex) {
                try { dbHolder[0] = dbf.newDocumentBuilder(); dbHolder[0].setErrorHandler(null); }
                catch (Exception ignored) {}
                log("⚠  " + nome + ": " + ex.getClass().getSimpleName() + " — " + ex.getMessage());
            }
        }

        // ── ZIPs: cada entrada processada e descartada — sem guardar bytes ────
        for (Path zip : zips) {
            if (cancelado) { log("⛔ Cancelado pelo usuário."); flushLog(); return; }
            log("📦  Abrindo: " + zip.getFileName());
            flushLog();
            int novos = 0, duplic = 0;
            try (ZipFile zf = new ZipFile(zip.toFile())) {
                Enumeration<? extends ZipEntry> ents = zf.entries();
                while (ents.hasMoreElements()) {
                    if (cancelado) break;
                    ZipEntry entry = ents.nextElement();
                    if (!entry.isDirectory()) {
                        String nome = Paths.get(entry.getName()).getFileName().toString().toUpperCase();
                        if (nome.endsWith(".XML")) {
                            if (!nomesVistos.contains(nome)) {
                                nomesVistos.add(nome);
                                total[0]++;
                                cnt[0]++;
                                atualizarProgresso(cnt[0], total[0]);
                                if (cnt[0] % 200 == 0) {
                                    log("⏳ " + cnt[0] + "/" + total[0] + " arquivos — S-1210: " + s1210[0]);
                                    flushLog();
                                }
                                try (InputStream is = zf.getInputStream(entry)) {
                                    List<Registro> regs = parsearStream(nome, is, dbHolder[0]);
                                    if (!regs.isEmpty()) { s1210[0]++; fundirNoMapa(regs, mapa, nome); }
                                } catch (Throwable ex) {
                                    try { dbHolder[0] = dbf.newDocumentBuilder(); dbHolder[0].setErrorHandler(null); }
                                    catch (Exception ignored) {}
                                    log("⚠  " + nome + ": " + ex.getClass().getSimpleName() + " — " + ex.getMessage());
                                }
                                novos++; xmlsDeZip++;
                            } else {
                                duplic++;
                            }
                        }
                    }
                }
                log("    └─ " + novos + " XML(s) adicionado(s)"
                    + (duplic > 0 ? ", " + duplic + " duplicado(s) ignorado(s)" : ""));
            } catch (Exception ex) {
                log("⚠  ZIP " + zip.getFileName() + ": " + ex.getMessage());
            }
            flushLog();
        }
        if (cancelado) { log("⛔ Cancelado pelo usuário."); flushLog(); return; }

        // Filtro de divergências
        boolean filtrarDiverg = chkDivergencias.isSelected();
        if (filtrarDiverg) {
            mapa.entrySet().removeIf(e -> !isDivergencia(e.getValue()));
        }

        log("─────────────────────────────────────");
        log("XMLs diretos       : " + xmlDiretos.size());
        log("ZIPs encontrados   : " + zips.size());
        log("XMLs extraídos ZIP : " + xmlsDeZip);
        log("S-1210 encontrados : " + s1210[0]);
        if (filtrarDiverg) log("Modo               : apenas divergências");
        log("Linhas na planilha : " + mapa.size());
        flushLog(); // garante que todo o log chega à UI antes dos diálogos

        if (s1210[0] == 0) {
            flushLog();
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(0);
                progressBar.setString("Nenhum S-1210");
                dlgAviso("Aviso",
                    "Nenhum arquivo S-1210 foi encontrado na pasta selecionada.\n\n" +
                    "Verifique se os arquivos possuem \"S-1210\" ou \"1210\" no nome.");
            });
            return;
        }

        if (filtrarDiverg && mapa.isEmpty()) {
            flushLog();
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(0);
                progressBar.setString("Sem divergências");
                dlgInfo("Sem divergências",
                    "Nenhuma divergência encontrada nos arquivos processados.\n\n" +
                    "Todos os períodos de apuração coincidem com os períodos de referência.");
            });
            return;
        }

        gerarExcel(new ArrayList<>(mapa.values()), saidaCaminho);

        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(100);
            progressBar.setString("Concluído!");
            int resp = dlgConfirmar("Concluído",
                "Planilha gerada com sucesso!\n\n" + saidaCaminho +
                "\n\nDeseja abrir o arquivo agora?");
            if (resp == JOptionPane.YES_OPTION) {
                try { Desktop.getDesktop().open(new File(saidaCaminho)); }
                catch (Exception ex) { dlgErro("Erro", "Não foi possível abrir o arquivo:\n" + ex.getMessage()); }
            }
        });
    }

    /** Atualiza a barra de progresso apenas quando o percentual muda. */
    private void atualizarProgresso(int atual, int total) {
        int pct = total > 0 ? (int)(atual * 100.0 / total) : 0;
        if (pct == ultimoPctBar) return;
        ultimoPctBar = pct;
        SwingUtilities.invokeLater(() -> {
            progressBar.setValue(pct);
            progressBar.setString(pct + "%");
        });
    }

    /**
     * Insere/substitui registros no mapa global.
     * Chave: CPF|perApur — se já existir entrada para o mesmo CPF+perApur,
     * prevalece o registro com dhProcessamento mais recente (comparação lexicográfica
     * em ISO-8601 é equivalente à cronológica). Ambos os arquivos ficam registrados
     * na coluna "Arquivo Origem" para rastreabilidade.
     */
    private void fundirNoMapa(List<Registro> registros, Map<String, Registro> mapa, String nomeExibicao) {
        for (Registro r : registros) {
            r.adicionarOrigem(nomeExibicao);
            String chave = r.cpf + "|" + r.perApur;
            Registro ex  = mapa.get(chave);
            if (ex == null) {
                mapa.put(chave, r);
            } else {
                // Mesmo CPF+perApur: decide pelo dhProcessamento mais recente
                boolean novoEhMaisRecente =
                    r.dhProcessamento.compareTo(ex.dhProcessamento) > 0;

                if (novoEhMaisRecente) {
                    // Novo prevalece — preserva origens do antigo para rastreabilidade
                    ex.origens.forEach(r::adicionarOrigem);
                    log("🔄 Substituição — CPF " + r.cpf
                        + " | perApur " + r.perApur
                        + " | novo: " + r.dhProcessamento
                            + " (" + nomeExibicao + ")"
                        + " | substituiu: " + ex.dhProcessamento
                            + " (" + ex.origensStr() + ")");
                    mapa.put(chave, r);
                } else {
                    // Existente prevalece — registra origem do descartado
                    r.origens.forEach(ex::adicionarOrigem);
                    log("🔄 Descartado — CPF " + r.cpf
                        + " | perApur " + r.perApur
                        + " | descartado: " + r.dhProcessamento
                            + " (" + nomeExibicao + ")"
                        + " | prevalece: " + ex.dhProcessamento
                            + " (" + ex.origensStr() + ")");
                }
            }
        }
    }

    // =========================================================================
    // Parse de um arquivo XML a partir de InputStream (sem carregar bytes inteiros)
    // =========================================================================
    private List<Registro> parsearStream(String nomeUpper, InputStream rawIs,
                                         DocumentBuilder dbf) throws Exception {
        boolean ehS1210porNome = nomeUpper.contains(".S-1210.");

        // BufferedInputStream com buffer de 64 KB; mark/reset para detectar tipo sem reler
        java.io.BufferedInputStream bis = new java.io.BufferedInputStream(rawIs, 65536);

        if (!ehS1210porNome) {
            bis.mark(4096);
            byte[] trecho = new byte[4096];
            int lidos = bis.read(trecho);
            if (lidos > 0) {
                String s = new String(trecho, 0, lidos, java.nio.charset.StandardCharsets.UTF_8);
                if (!s.contains("evtPgtos") && !s.contains("infoPgto")) {
                    return Collections.emptyList();
                }
            }
            bis.reset();
        }

        Document doc = dbf.parse(bis); // dbf aqui já é o DocumentBuilder reutilizável
        doc.getDocumentElement().normalize();

        String perApur = primeiroTexto(doc, "perApur");
        if (perApur == null || perApur.isBlank()) perApur = "(sem perApur)";

        String nrRecibo        = primeiroTexto(doc, "nrRecibo");
        String dhProcessamento = primeiroTexto(doc, "dhProcessamento");
        if (dhProcessamento == null) dhProcessamento = "";

        NodeList benefs = doc.getElementsByTagName("ideBenef");
        if (benefs.getLength() == 0)
            benefs = doc.getElementsByTagName("infoBenef");

        List<Registro> resultado = new ArrayList<>();

        for (int i = 0; i < benefs.getLength(); i++) {
            Element benef = (Element) benefs.item(i);
            String cpf = textoFilho(benef, "cpfBenef");
            if (cpf == null || cpf.isBlank()) continue;

            final String perApurFinal = perApur;
            Registro reg = resultado.stream()
                .filter(r -> r.cpf.equals(cpf) && r.perApur.equals(perApurFinal))
                .findFirst().orElse(null);
            if (reg == null) {
                reg = new Registro();
                reg.cpf            = cpf;
                reg.perApur        = perApur;
                reg.nrRecibo       = nrRecibo;
                reg.dhProcessamento = dhProcessamento;
                resultado.add(reg);
            }

            NodeList pagamentos = benef.getElementsByTagName("infoPgto");
            for (int j = 0; j < pagamentos.getLength(); j++) {
                Element pag = (Element) pagamentos.item(j);
                reg.adicionarDmDev(textoFilho(pag, "ideDmDev"));
                reg.adicionarPerRef(textoFilho(pag, "perRef"));
            }
        }

        return resultado;
    }

    // =========================================================================
    // Verifica divergência: ao menos um perRef é diferente de perApur.
    // Exceção: perRef anual "YYYY" (ex.: 2025) é considerado igual a "YYYY-12".
    // =========================================================================
    private boolean isDivergencia(Registro r) {
        for (String ref : r.perRefs) {
            if (r.perApur.equals(ref)) continue;
            // Referência anual: "2025" coincide com perApur "2025-12"
            if (!ref.contains("-") && r.perApur.equals(ref + "-12")) continue;
            // Encontrou ao menos um perRef que não bate com perApur
            return true;
        }
        return false;
    }

    // =========================================================================
    // Geração da planilha Excel
    // =========================================================================
    private void gerarExcel(List<Registro> registros, String caminho) throws Exception {
        log("Gerando planilha Excel...");

        registros.sort(Comparator
            .comparing((Registro r) -> r.perApur)
            .thenComparing(r -> r.cpf)
            .thenComparing(r -> r.nrRecibo != null ? r.nrRecibo : ""));

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet aba = wb.createSheet("S-1210");
            aba.createFreezePane(0, 1);

            CellStyle estHeader = wb.createCellStyle();
            org.apache.poi.ss.usermodel.Font fHeader = wb.createFont();
            fHeader.setBold(true);
            fHeader.setColor(IndexedColors.WHITE.getIndex());
            fHeader.setFontHeightInPoints((short) 11);
            estHeader.setFont(fHeader);
            estHeader.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            estHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estHeader.setBorderBottom(BorderStyle.THIN);
            estHeader.setBottomBorderColor(IndexedColors.LIGHT_BLUE.getIndex());
            estHeader.setAlignment(HorizontalAlignment.CENTER);

            CellStyle estPar = wb.createCellStyle();
            estPar.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            estPar.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle estImpar = wb.createCellStyle();

            String[] cols = {"CPF", "ideDmDev", "Período de Apuração", "Período de Referência", "Número Recibo", "Arquivo Origem"};
            Row header = aba.createRow(0);
            header.setHeightInPoints(22);
            for (int c = 0; c < cols.length; c++) {
                Cell cell = header.createCell(c);
                cell.setCellValue(cols[c]);
                cell.setCellStyle(estHeader);
            }

            int linha = 1;
            for (Registro r : registros) {
                Row row = aba.createRow(linha);
                row.setHeightInPoints(16);
                CellStyle est = (linha % 2 == 0) ? estPar : estImpar;
                String[] vals = {
                    r.cpf,
                    r.idmDevsStr(),
                    r.perApur,
                    r.perRefsStr(),
                    r.nrRecibo != null ? r.nrRecibo : "",
                    r.origensStr()
                };
                for (int c = 0; c < vals.length; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue(vals[c]);
                    cell.setCellStyle(est);
                }
                linha++;
            }

            int[] larguras = {4500, 14000, 5000, 5500, 8000, 16000};
            for (int c = 0; c < larguras.length; c++) aba.setColumnWidth(c, larguras[c]);

            try (FileOutputStream fos = new FileOutputStream(caminho)) { wb.write(fos); }
        }

        log("Planilha salva: " + caminho);
    }

    // ── Utilitários XML ───────────────────────────────────────────────────────
    private String textoFilho(Element pai, String tag) {
        NodeList nl = pai.getElementsByTagName(tag);
        if (nl.getLength() > 0) {
            String t = nl.item(0).getTextContent();
            return (t != null && !t.isBlank()) ? t.trim() : null;
        }
        return null;
    }

    private String primeiroTexto(Document doc, String tag) {
        NodeList nl = doc.getElementsByTagName(tag);
        for (int i = 0; i < nl.getLength(); i++) {
            String t = nl.item(i).getTextContent();
            if (t != null && !t.isBlank()) return t.trim();
        }
        return null;
    }

    /** Enfileira uma linha de log — sem tocar na EDT imediatamente. */
    private void log(String msg) {
        logBuffer.add(msg);
    }

    /**
     * Drena o buffer de log para o JTextArea em um único invokeLater.
     * Chamar periodicamente durante o processamento (ex: a cada N arquivos).
     */
    private void flushLog() {
        if (logBuffer.isEmpty()) return;
        StringBuilder sb = new StringBuilder();
        String linha;
        while ((linha = logBuffer.poll()) != null) sb.append(linha).append('\n');
        final String bloco = sb.toString();
        SwingUtilities.invokeLater(() -> {
            txtLog.append(bloco);
            txtLog.setCaretPosition(txtLog.getDocument().getLength());
        });
    }

    /** Grava stack trace completo em s1210_erro.log para diagnóstico. */
    private static void gravarLogErro(Throwable ex) {
        try {
            java.io.File f = new java.io.File(
                System.getProperty("user.home"), "s1210_erro.log");
            try (java.io.PrintWriter pw = new java.io.PrintWriter(
                    new java.io.FileWriter(f, true))) {
                pw.println("=== " + new java.util.Date() + " ===");
                ex.printStackTrace(pw);
                pw.println();
            }
        } catch (Exception ignored) {}
    }

    // =========================================================================
    public static void main(String[] args) {
        // Handler global — garante que qualquer crash inesperado grava log
        Thread.setDefaultUncaughtExceptionHandler((t, ex) -> {
            gravarLogErro(ex);
            ex.printStackTrace();
        });

        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); }
        catch (Exception ignored) {}
        SwingUtilities.invokeLater(() -> new S1210Extrator().setVisible(true));
    }
}
