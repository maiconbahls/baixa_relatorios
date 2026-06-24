module.exports = async (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
    res.setHeader(
        'Access-Control-Allow-Headers',
        'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
    );

    // Handle OPTIONS request pre-flight
    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    // Force POST
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Método não permitido. Utilize POST.' });
    }

    const { usuario, senha, id } = req.body || {};

    if (!usuario || !senha || !id) {
        return res.status(400).json({ error: 'Dados obrigatórios ausentes (usuario, senha ou id).' });
    }

    const domain = process.env.REPORT_SYSTEM_DOMAIN;
    if (!domain) {
        return res.status(500).json({ error: 'Configuração do servidor ausente: a variável REPORT_SYSTEM_DOMAIN não foi definida.' });
    }

    try {
        // 1. Efetuar Login no Reportview
        const loginUrl = `http://${domain}/login.php`;
        const params = new URLSearchParams();
        params.append('usuario', usuario);
        params.append('senha', senha);

        const loginResponse = await fetch(loginUrl, {
            method: 'POST',
            body: params,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            },
            redirect: 'manual'
        });

        const setCookieHeader = loginResponse.headers.get('set-cookie');
        if (!setCookieHeader) {
            return res.status(401).json({ error: 'Falha na autenticação do Interview. Verifique usuário e senha.' });
        }

        // Extrai o cookie da sessão PHPSESSID
        const cookie = setCookieHeader.split(';')[0];

        // 2. Obter dados do Relatório
        const dataUrl = `http://${domain}/relatorio/dados.php?id=${id}`;
        const dataResponse = await fetch(dataUrl, {
            headers: {
                'Cookie': cookie,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
        });

        if (dataResponse.status !== 200) {
            return res.status(dataResponse.status).json({ 
                error: `O servidor do Interview retornou o status de erro: ${dataResponse.status}` 
            });
        }

        const text = await dataResponse.text();

        // Se o resultado contiver a estrutura da página de login, significa que as credenciais falharam
        if (text.includes('Login - Interview') || text.trim().startsWith('<!doctype html')) {
            return res.status(401).json({ 
                error: 'Falha no login do Interview ou acesso negado ao relatório. Verifique suas credenciais.' 
            });
        }

        try {
            // Remove o UTF-8 BOM se presente na resposta do PHP
            const cleanText = text.replace(/^\uFEFF/, '');
            const jsonData = JSON.parse(cleanText);
            
            // Retorna os dados originais formatados para o front-end
            return res.status(200).json(jsonData);
        } catch (parseError) {
            return res.status(500).json({ 
                error: 'Falha ao processar os dados do relatório. O servidor não retornou um JSON válido.',
                preview: text.substring(0, 150)
            });
        }

    } catch (error) {
        return res.status(500).json({ error: `Erro de conexão interna no servidor proxy: ${error.message}` });
    }
};
