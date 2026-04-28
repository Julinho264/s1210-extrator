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
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

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
    private JTextArea   txtLog;
    private JProgressBar progressBar;
    private JLabel      lblStatus;
    private JCheckBox   chkDivergencias;

    // =========================================================================
    public S1210Extrator() {
        super("SIGEP — Extrator S-1210");
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

        JPanel actionRow = new JPanel(new BorderLayout(10, 0));
        actionRow.setOpaque(false);
        actionRow.add(progressBar, BorderLayout.CENTER);
        actionRow.add(btnGerar,    BorderLayout.EAST);

        lblStatus = new JLabel("© SIGEP — Sistema Integrado de Gestão Pública");
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

    // ── Seleção de caminhos ───────────────────────────────────────────────────
    private void selecionarPasta(ActionEvent e) {
        JFileChooser fc = new JFileChooser(System.getProperty("user.home"));
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        fc.setDialogTitle("Selecionar pasta com XMLs S-1210");
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File pasta = fc.getSelectedFile();
            txtPasta.setText(pasta.getAbsolutePath());
            if (txtSaida.getText().isBlank())
                txtSaida.setText(new File(pasta, "s1210_resultado.xlsx").getAbsolutePath());
        }
    }

    private void selecionarSaida(ActionEvent e) {
        JFileChooser fc = new JFileChooser(System.getProperty("user.home"));
        fc.setDialogTitle("Salvar planilha como...");
        fc.setSelectedFile(new File("s1210_resultado.xlsx"));
        if (fc.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            String caminho = fc.getSelectedFile().getAbsolutePath();
            if (!caminho.toLowerCase().endsWith(".xlsx")) caminho += ".xlsx";
            txtSaida.setText(caminho);
        }
    }

    // ── Início do processamento (thread separada) ─────────────────────────────
    private void iniciarGeracao() {
        String pasta = txtPasta.getText().trim();
        String saida = txtSaida.getText().trim();

        if (pasta.isBlank() || saida.isBlank()) {
            dlgAviso("Atenção", "Selecione a pasta de entrada e o arquivo de saída.");
            return;
        }

        btnGerar.setEnabled(false);
        txtLog.setText("");
        progressBar.setValue(0);
        progressBar.setString("Processando...");

        new Thread(() -> {
            try {
                processar(pasta, saida);
            } catch (Exception ex) {
                log("ERRO: " + ex.getMessage());
                SwingUtilities.invokeLater(() ->
                    dlgErro("Erro", "Erro durante o processamento:\n" + ex.getMessage()));
            } finally {
                SwingUtilities.invokeLater(() -> btnGerar.setEnabled(true));
            }
        }).start();
    }

    // =========================================================================
    // Estrutura de dados: CPF + perApur → [ideDmDev], [perRef], nrRecibo
    //   perApur  — de <ideEvento><perApur>, único por arquivo
    //   perRef   — de cada <infoPgto><perRef>, sem repetição por beneficiário
    //   nrRecibo — de <retornoEvento><recibo><nrRecibo>
    // =========================================================================
    private static class Registro {
        String cpf;
        String perApur;
        String nrRecibo;
        LinkedHashSet<String> idmDevs  = new LinkedHashSet<>();
        LinkedHashSet<String> perRefs  = new LinkedHashSet<>();

        void adicionarDmDev(String v) {
            if (v != null && !v.isBlank()) idmDevs.add(v.trim());
        }

        void adicionarPerRef(String v) {
            if (v != null && !v.isBlank()) perRefs.add(v.trim());
        }

        String idmDevsStr()  { return String.join("; ", idmDevs); }
        String perRefsStr()  { return String.join("; ", perRefs); }
    }

    // =========================================================================
    // Processamento principal
    // =========================================================================
    private void processar(String pastaCaminho, String saidaCaminho) throws Exception {

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

        // ── Mapa nome→bytes (chave em maiúsculas para deduplicação) ──────────
        // Arquivos diretos têm prioridade; XMLs de ZIP são ignorados se o nome
        // já estiver presente como arquivo direto.
        Map<String, byte[]> fontes = new LinkedHashMap<>();

        for (Path xml : xmlDiretos) {
            fontes.put(xml.getFileName().toString().toUpperCase(), Files.readAllBytes(xml));
        }

        int xmlsDeZip = 0;
        for (Path zip : zips) {
            try (ZipInputStream zis = new ZipInputStream(Files.newInputStream(zip))) {
                ZipEntry entry;
                while ((entry = zis.getNextEntry()) != null) {
                    if (!entry.isDirectory()) {
                        String nome = Paths.get(entry.getName()).getFileName().toString();
                        if (nome.toLowerCase().endsWith(".xml")
                                && !fontes.containsKey(nome.toUpperCase())) {
                            fontes.put(nome.toUpperCase(), zis.readAllBytes());
                            xmlsDeZip++;
                        }
                    }
                    zis.closeEntry();
                }
            } catch (Exception ex) {
                log("⚠  ZIP " + zip.getFileName() + ": " + ex.getMessage());
            }
        }

        log("XMLs diretos: " + xmlDiretos.size()
            + " | ZIPs: " + zips.size()
            + " | XMLs extraídos de ZIPs: " + xmlsDeZip
            + " | Total único: " + fontes.size());

        Map<String, Registro> mapa = new LinkedHashMap<>();
        int total = fontes.size();
        int contador = 0, s1210Encontrados = 0;

        for (Map.Entry<String, byte[]> fonte : fontes.entrySet()) {
            contador++;
            final int pct = (int)(contador * 100.0 / total);
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(pct);
                progressBar.setString(pct + "%");
            });

            String nomeExibicao = fonte.getKey();
            try {
                List<Registro> registros = parsearArquivo(nomeExibicao, fonte.getValue());
                if (!registros.isEmpty()) {
                    s1210Encontrados++;
                    for (Registro r : registros) {
                        String recibo = (r.nrRecibo != null && !r.nrRecibo.isBlank()) ? r.nrRecibo : "";
                        String chave  = r.cpf + "|" + r.perApur + "|" + recibo;
                        Registro ex = mapa.get(chave);
                        if (ex == null) {
                            mapa.put(chave, r);
                        } else {
                            r.idmDevs.forEach(ex::adicionarDmDev);
                            r.perRefs.forEach(ex::adicionarPerRef);
                        }
                    }
                    long nBenef = registros.stream().map(r -> r.cpf).distinct().count();
                    log("✔  " + nomeExibicao + " — " + nBenef + " beneficiário(s)");
                }
            } catch (Exception ex) {
                log("⚠  " + nomeExibicao + ": " + ex.getMessage());
            }
        }

        // Filtro de divergências (aplica antes de contar as linhas finais)
        boolean filtrarDiverg = chkDivergencias.isSelected();
        if (filtrarDiverg) {
            mapa.entrySet().removeIf(e -> !isDivergencia(e.getValue()));
        }

        log("─────────────────────────────────────");
        log("Arquivos S-1210 encontrados : " + s1210Encontrados);
        if (filtrarDiverg) {
            log("Modo: apenas divergências");
        }
        log("Linhas na planilha (CPF+período): " + mapa.size());

        if (s1210Encontrados == 0) {
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
                try {
                    Desktop.getDesktop().open(new File(saidaCaminho));
                } catch (Exception ex) {
                    dlgErro("Erro", "Não foi possível abrir o arquivo:\n" + ex.getMessage());
                }
            }
        });
    }

    // =========================================================================
    // Parse de um arquivo XML (nome em maiúsculas, bytes já lidos)
    // =========================================================================
    private List<Registro> parsearArquivo(String nomeUpper, byte[] bytes) throws Exception {
        boolean ehS1210porNome = nomeUpper.contains(".S-1210.");

        if (!ehS1210porNome) {
            String trecho = new String(bytes, 0, Math.min(bytes.length, 4096), "UTF-8");
            if (!trecho.contains("evtPgtos") && !trecho.contains("infoPgto")) {
                return Collections.emptyList();
            }
        }

        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(false);
        DocumentBuilder db = dbf.newDocumentBuilder();
        db.setErrorHandler(null);
        Document doc = db.parse(new java.io.ByteArrayInputStream(bytes));
        doc.getDocumentElement().normalize();

        String perApur = primeiroTexto(doc, "perApur");
        if (perApur == null || perApur.isBlank()) perApur = "(sem perApur)";

        String nrRecibo = primeiroTexto(doc, "nrRecibo");

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
                reg.cpf      = cpf;
                reg.perApur  = perApur;
                reg.nrRecibo = nrRecibo;
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

            String[] cols = {"CPF", "ideDmDev", "Período de Apuração", "Período de Referência", "Número Recibo"};
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
                    r.nrRecibo != null ? r.nrRecibo : ""
                };
                for (int c = 0; c < vals.length; c++) {
                    Cell cell = row.createCell(c);
                    cell.setCellValue(vals[c]);
                    cell.setCellStyle(est);
                }
                linha++;
            }

            int[] larguras = {4500, 14000, 5000, 5500, 8000};
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

    private void log(String msg) {
        SwingUtilities.invokeLater(() -> {
            txtLog.append(msg + "\n");
            txtLog.setCaretPosition(txtLog.getDocument().getLength());
        });
    }

    // =========================================================================
    public static void main(String[] args) {
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); }
        catch (Exception ignored) {}
        SwingUtilities.invokeLater(() -> new S1210Extrator().setVisible(true));
    }
}
