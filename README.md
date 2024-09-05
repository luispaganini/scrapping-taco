# Projeto C# de Web Scraping com EF Core

Este projeto é uma aplicação C# que realiza scraping de dados de alimentos e informações nutricionais de uma página web e armazena esses dados em um banco de dados SQL Server.

## Requisitos

Antes de começar, certifique-se de ter o seguinte instalado em seu sistema:

- [.NET SDK](https://dotnet.microsoft.com/download) (versão 6.0 ou superior)
- [SQL Server](https://www.microsoft.com/en-us/sql-server/sql-server-downloads) (ou [SQL Server Express](https://www.microsoft.com/en-us/sql-server/sql-server-editions))
- [SQL Server Management Studio (SSMS)](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms) (opcional, para gerenciamento do banco de dados)

## Instalação

1. Clone este repositório para sua máquina local:

   ```bash
   git clone https://github.com/seu-usuario/seu-projeto.git
   ```

2. Acesse o diretório do projeto:

   ```bash
   cd seu-projeto
   ```

3. Restaure as dependências do projeto:

   ```bash
   dotnet restore
   ```

4. Configure o banco de dados:

   - **Crie o Banco de Dados SQL Server:**

     - Abra o [SQL Server Management Studio (SSMS)](https://docs.microsoft.com/en-us/sql/ssms/download-sql-server-management-studio-ssms) e conecte-se ao seu servidor SQL Server.
     - No painel "Object Explorer", clique com o botão direito em "Databases" e selecione "New Database...".
     - Dê um nome ao banco de dados, por exemplo, `Balanced_Life`, e clique em "OK".

   - **Configure a String de Conexão:**

     - No código fonte, a string de conexão está configurada diretamente no `Program.cs`:

       ```csharp
       var optionsBuilder = new DbContextOptionsBuilder<Balanced_LifeContext>();
       optionsBuilder.UseSqlServer("Data Source=localhost\\MSSQLSERVER02;Initial Catalog=Balanced_Life;Integrated Security=True;Encrypt=True;TrustServerCertificate=YES");
       ```

     - Certifique-se de que os detalhes da string de conexão (`Data Source`, `Initial Catalog`, etc.) estão corretos para o seu ambiente.


## Executar a Aplicação

Para rodar a aplicação, utilize o seguinte comando a partir do diretório raiz do projeto:

```bash
dotnet run
```

A aplicação irá:

1. Realizar scraping das páginas web especificadas para coletar dados sobre alimentos e suas informações nutricionais.
2. Armazenar essas informações no banco de dados SQL Server configurado.

## Estrutura de Pastas

```bash
├── Program.cs          # Ponto de entrada da aplicação
├── Models              # Modelos de dados
│   └── SomeModel.cs    # Modelos utilizados no scraping
│   └── Balanced_LifeContext.cs # Contexto do banco de dados (Entity Framework)
├── appsettings.json    # Configurações da aplicação (se aplicável)
├── .gitignore          # Arquivos e pastas a serem ignorados pelo Git
└── seu-projeto.sln     # Solução do projeto
```

## Scripts Disponíveis

- `run`: Executa a aplicação para realizar o scraping e armazenar dados no banco de dados.

## Contribuição

Sinta-se à vontade para abrir uma issue ou pull request para sugestões de melhorias ou correções.

---

Siga este guia para configurar e executar sua aplicação C# para scraping de dados e armazenamento em um banco de dados SQL Server. Para mais informações, consulte a documentação oficial do .NET e do Entity Framework.
