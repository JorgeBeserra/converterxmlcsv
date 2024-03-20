use clap::{Arg, ArgAction, Command};
use serde::{Deserialize, Serialize};
use serde_xml_rs::from_reader;
use std::fs::File;
use std::io::{BufReader, BufWriter, stdin};
use std::error::Error;
use csv::WriterBuilder;
use glob::glob;
use dialoguer::Select;
use std::path::{Path, PathBuf};
use colored::*;

#[derive(Debug, Deserialize)]
struct Funcionario {
    #[serde(rename = "CPF")]
    cpf: String,
    #[serde(rename = "Valor")]
    valor: String,
    #[serde(rename = "MetaPremio", default)]
    meta_premio: Option<String>,
}

#[derive(Debug, Deserialize)]
struct Empresa {
    #[serde(rename = "Fantasia")]
    fantasia: String,
    #[serde(rename = "Razao")]
    razao: String,
    #[serde(rename = "CNPJ")]
    cnpj: String,
    #[serde(rename = "MesAno")]
    mes_ano: String,
    #[serde(rename = "Funcionario")]
    funcionarios: Option<Vec<Funcionario>>,
}

#[derive(Debug, Deserialize)]
enum TipoArquivo {
    Comissao(Comissao),
    Vales(Vales),
}

#[derive(Debug, Deserialize)]
struct Comissao {
    #[serde(rename = "Empresa")]
    empresa: Empresa,
}

#[derive(Debug, Deserialize)]
struct Vales {
    #[serde(rename = "Empresa")]
    empresa: Empresa,
}

fn main() -> Result<(), Box<dyn Error>> {
    let matches = clap::Command::new("Conversor XML para CSV")
        .version("0.1.0")
        .author("Jorge Beserra <jorgebeserra@gmail.com>")
        .about("Converte arquivos XML de comissões ou vales para CSV")
        .get_matches();

    // Mensagem de boas-vindas
    println!("{}", "Bem-vindo ao Conversor XML para CSV!".bright_green());
    println!("{}", "Desenvolvido por Jorge Beserra <jorgebeserra@gmail.com>".bright_yellow());
    println!("{}", "Repositório no GitHub: https://github.com/jorgebeserra/conversorxmlcsv\n".bright_yellow());

    let mut xml_files: Vec<PathBuf> = glob("*.xml")?
        .filter_map(Result::ok)
        .collect();  

    if xml_files.is_empty() {
        println!("{}", "Não foram encontrados arquivos XML na pasta.".bright_red());
        return Ok(());
    }

    let selection = Select::new()
        .items(&xml_files.iter().map(|path| path.file_name().unwrap().to_str().unwrap()).collect::<Vec<&str>>())
        .default(0)
        .with_prompt("Escolha o arquivo XML a ser convertido:")
        .interact()?;

    let selected_file = &xml_files[selection];
    let selected_file_stem = selected_file.file_stem().unwrap().to_str().unwrap();


    let file = File::open(selected_file)?;
    let reader = BufReader::new(file);

    let tipo_arquivo: TipoArquivo = match selected_file_stem.split('_').next() {
        Some("comissao") => {
            let comissao: Comissao = serde_xml_rs::from_reader(reader)?;
            TipoArquivo::Comissao(comissao)
        }
        Some("vales") => {
            let vale: Vales = serde_xml_rs::from_reader(reader)?;
            TipoArquivo::Vales(vale)
        }
        _ => return Err("Tipo de arquivo não suportado.".into()),
    };

    match tipo_arquivo {
        TipoArquivo::Comissao(comissao) => handle_comissao(comissao, selected_file)?,
        TipoArquivo::Vales(vale) => handle_vale(vale, selected_file)?,
    }

    println!("{}", "Pressione Enter para sair...".bright_cyan());
    let _ = stdin().read_line(&mut String::new());

    Ok(())
}

fn handle_comissao(comissao: Comissao, selected_file: &PathBuf) -> Result<(), Box<dyn Error>> {
    handle_arquivo_comissao(comissao.empresa, selected_file)
}

fn handle_vale(vale: Vales, selected_file: &PathBuf) -> Result<(), Box<dyn Error>> {
    handle_arquivo_vales(vale.empresa, selected_file)
}

fn handle_arquivo_comissao(empresa: Empresa, selected_file: &PathBuf) -> Result<(), Box<dyn Error>> {

        // Verifica se a empresa possui funcionários
        let funcionarios = match empresa.funcionarios {
            Some(funcionarios) => funcionarios,
            None => {
                println!("{}", "O arquivo XML não contém funcionários. Nenhum dado será exportado para o CSV.".bright_yellow());
                return Ok(());
            }
        };

    let csv_file_path = Path::new(selected_file).with_extension("csv");
    let csv_file = File::create(&csv_file_path)?;
    
    let mut csv_writer = csv::WriterBuilder::new()
        .delimiter(b';')
        .from_writer(csv_file);

    let mut total_comissao: f64 = 0.0;
    let mut total_meta: f64 = 0.0;
    let mut quantidade_funcionarios = 0;

    // Escreve o cabeçalho no arquivo CSV
    csv_writer.write_record(&["Fantasia", "Razao", "CNPJ", "MesAno", "CPF", "Valor", "MetaPremio"])?;

    // Itera sobre os funcionários da empresa e escreve seus dados no arquivo CSV
    for funcionario in funcionarios {
        let meta_premio = if let Some(meta_premio) = &funcionario.meta_premio {
            meta_premio
        } else {
            ""
        };

        total_comissao += funcionario.valor.parse::<f64>().unwrap_or(0.0);
        total_meta += meta_premio.parse::<f64>().unwrap_or(0.0);
        quantidade_funcionarios += 1;
        
        csv_writer.write_record(&[
            &empresa.fantasia,
            &empresa.razao,
            &empresa.cnpj,
            &empresa.mes_ano,
            &funcionario.cpf,
            &funcionario.valor,
            meta_premio
        ])?;
    }

    csv_writer.flush()?;

    println!("{}", format!("Dados exportados para {} com sucesso! \nQuantidade de funcionários: {}. \nTotal de comissão: R$ {:.2}\nTotal por meta: R$ {:.2}", csv_file_path.display(), quantidade_funcionarios, total_comissao, total_meta).bright_green());

    Ok(())
}

fn handle_arquivo_vales(empresa: Empresa, selected_file: &PathBuf) -> Result<(), Box<dyn Error>> {
    
    // Verifica se a empresa possui funcionários
    let funcionarios = match empresa.funcionarios {
        Some(funcionarios) => funcionarios,
        None => {
            println!("{}", "O arquivo XML não contém funcionários. Nenhum dado será exportado para o CSV.".bright_yellow());
            return Ok(());
        }
    };

    let csv_file_path = Path::new(selected_file).with_extension("csv");
    let csv_file = File::create(&csv_file_path)?;
    
    let mut csv_writer = csv::WriterBuilder::new()
        .delimiter(b';')
        .from_writer(csv_file);

    // Escreve o cabeçalho no arquivo CSV
    csv_writer.write_record(&["Fantasia", "Razao", "CNPJ", "MesAno", "CPF", "Valor"])?;

    let mut total_vales: f64 = 0.0;
    let mut quantidade_funcionarios = 0;

    // Itera sobre os funcionários da empresa e escreve seus dados no arquivo CSV
    for funcionario in funcionarios {
        let meta_premio = if let Some(meta_premio) = &funcionario.meta_premio {
            meta_premio
        } else {
            ""
        };

        total_vales += funcionario.valor.parse::<f64>().unwrap_or(0.0);
        quantidade_funcionarios += 1;
        
        csv_writer.write_record(&[
            &empresa.fantasia,
            &empresa.razao,
            &empresa.cnpj,
            &empresa.mes_ano,
            &funcionario.cpf,
            &funcionario.valor
        ])?;
    }

    csv_writer.flush()?;

    println!("{}", format!("Dados exportados para {} com sucesso! \nQuantidade de funcionários: {}. \nTotal de vales: R$ {:.2}", csv_file_path.display(), quantidade_funcionarios, total_vales).bright_green());

    Ok(())
}