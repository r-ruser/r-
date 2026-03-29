library(shiny)
library(rentrez)
library(openxlsx)
library(dplyr)
library(rvest)
library(readr)
library(readxl)
library(writexl)
library(httr)

options(shiny.maxRequestSize = 100 * 1024^2)

safe_read_html <- function(url, timeout_sec = 60) {
  tryCatch(
    {
      httr::GET(url, httr::timeout(timeout_sec)) |>
        httr::content(as = "text", encoding = "UTF-8") |>
        xml2::read_html()
    },
    error = function(e) NULL
  )
}

extract_abstract_from_pubmed <- function(pmid, timeout_sec = 60) {
  url <- paste0("https://pubmed.ncbi.nlm.nih.gov/", pmid)
  page <- safe_read_html(url, timeout_sec = timeout_sec)

  if (is.null(page)) return(NA_character_)

  eng_abstract <- tryCatch(
    rvest::html_text(rvest::html_elements(page, "#eng-abstract > p"), trim = TRUE),
    error = function(e) character(0)
  )
  abstract <- tryCatch(
    rvest::html_text(rvest::html_elements(page, "#abstract > p"), trim = TRUE),
    error = function(e) character(0)
  )

  if (length(eng_abstract) > 0) {
    paste(eng_abstract, collapse = " ")
  } else if (length(abstract) > 0) {
    paste(abstract, collapse = " ")
  } else {
    NA_character_
  }
}

parse_keywords <- function(txt) {
  if (is.null(txt) || !nzchar(trimws(txt))) return(character(0))
  x <- unlist(strsplit(txt, "[\n,，;；]+"))
  x <- trimws(x)
  x[nzchar(x)]
}

build_pubmed_query <- function(keywords,
                               start_year = NULL,
                               end_year = NULL,
                               join_mode = c("AND", "OR"),
                               article_type = NULL,
                               language = NULL,
                               custom_query = NULL) {
  join_mode <- match.arg(join_mode)

  if (!is.null(custom_query) && nzchar(trimws(custom_query))) {
    query <- trimws(custom_query)
  } else {
    if (length(keywords) == 0) stop("请至少输入一个关键词，或直接输入自定义检索式。")
    quoted_keywords <- sprintf("(%s)", keywords)
    query <- paste(quoted_keywords, collapse = paste0(" ", join_mode, " "))
  }

  if (!is.null(start_year) && !is.na(start_year) && !is.null(end_year) && !is.na(end_year)) {
    if (start_year > end_year) stop("起始年份不能大于结束年份。")
    date_filter <- sprintf("(\"%s/01/01\"[Date - Publication] : \"%s/12/31\"[Date - Publication])", start_year, end_year)
    query <- paste(query, "AND", date_filter)
  }

  if (!is.null(article_type) && nzchar(article_type) && article_type != "不限") {
    query <- paste(query, "AND", sprintf("\"%s\"[Publication Type]", article_type))
  }

  if (!is.null(language) && nzchar(language) && language != "不限") {
    lang_term <- switch(language,
                        "英文" = "english[Language]",
                        "中文" = "chinese[Language]",
                        "其他" = "NOT (english[Language] OR chinese[Language])",
                        NULL)
    if (!is.null(lang_term)) {
      query <- paste(query, "AND", lang_term)
    }
  }

  query
}

fetch_pubmed_results <- function(query, retmax = 20, timeout_sec = 60, include_abstract = TRUE) {
  set_config(config(timeout = timeout_sec))

  search_result <- rentrez::entrez_search(
    db = "pubmed",
    term = query,
    retmax = retmax,
    use_history = FALSE
  )

  ids <- search_result$ids
  if (length(ids) == 0) {
    return(data.frame(
      PMID = character(0),
      标题 = character(0),
      摘要 = character(0),
      ISO = character(0),
      发表年份 = character(0),
      类型 = character(0),
      影响因子 = numeric(0),
      stringsAsFactors = FALSE
    ))
  }

  summaries <- rentrez::entrez_summary(db = "pubmed", id = ids)

  extract_safe <- function(obj, field) {
    tryCatch(sapply(obj, function(x) x[[field]]), error = function(e) rep(NA_character_, length(obj)))
  }

  titles <- extract_safe(summaries, "title")
  journals <- extract_safe(summaries, "source")
  years <- extract_safe(summaries, "pubdate")
  pmids <- extract_safe(summaries, "uid")
  types <- tryCatch(
    sapply(summaries, function(x) paste(x$pubtype, collapse = "; ")),
    error = function(e) rep(NA_character_, length(ids))
  )

  if (isTRUE(include_abstract)) {
    abstracts <- vapply(pmids, function(pmid) {
      extract_abstract_from_pubmed(pmid, timeout_sec = timeout_sec)
    }, character(1))
  } else {
    abstracts <- rep(NA_character_, length(pmids))
  }

  data.frame(
    PMID = pmids,
    标题 = titles,
    摘要 = abstracts,
    ISO = journals,
    发表年份 = years,
    类型 = types,
    影响因子 = NA_real_,
    stringsAsFactors = FALSE
  )
}

match_jcr_if <- function(pubmed_df, jcr_file, if_file) {
  out <- pubmed_df

  jcr_ext <- tools::file_ext(jcr_file$datapath)
  if (tolower(jcr_ext) %in% c("txt", "tsv")) {
    jcr_df <- readr::read_delim(jcr_file$datapath, delim = "\t", show_col_types = FALSE)
  } else {
    jcr_df <- readxl::read_excel(jcr_file$datapath)
  }

  if_df <- readxl::read_excel(if_file$datapath)

  needed_jcr_cols <- c("ISO", "JCR")
  needed_if_cols <- c("JCRAbbreviation", "JIF")

  if (!all(needed_jcr_cols %in% names(jcr_df))) {
    stop("JCR 对照文件缺少必要列：ISO, JCR")
  }
  if (!all(needed_if_cols %in% names(if_df))) {
    stop("IF 文件缺少必要列：JCRAbbreviation, JIF")
  }

  lookup_jcr <- setNames(jcr_df$JCR, jcr_df$ISO)
  out$JCR <- lookup_jcr[out$ISO]

  lookup_if <- setNames(if_df$JIF, if_df$JCRAbbreviation)
  out$IF <- lookup_if[out$JCR]

  out
}

save_results_excel <- function(df, file_path, query_text) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "PubMed Results")
  openxlsx::writeData(wb, "PubMed Results", df)
  openxlsx::freezePane(wb, sheet = "PubMed Results", firstRow = TRUE)

  openxlsx::addWorksheet(wb, "Query")
  openxlsx::writeData(wb, "Query", data.frame(检索式 = query_text, stringsAsFactors = FALSE))

  openxlsx::saveWorkbook(wb, file = file_path, overwrite = TRUE)
}

ui <- fluidPage(
  titlePanel("PubMed 文献抓取可视化界面"),
  sidebarLayout(
    sidebarPanel(
      radioButtons(
        "query_mode",
        "检索方式",
        choices = c("关键词组合" = "keywords", "自定义检索式" = "custom"),
        selected = "keywords"
      ),
      conditionalPanel(
        condition = "input.query_mode == 'keywords'",
        textAreaInput(
          "keywords",
          "关键词（可换行、逗号或分号分隔）",
          placeholder = "例如：melanoma\nCTC",
          rows = 5
        ),
        radioButtons(
          "join_mode",
          "关键词连接方式",
          choices = c("AND" = "AND", "OR" = "OR"),
          selected = "AND",
          inline = TRUE
        )
      ),
      conditionalPanel(
        condition = "input.query_mode == 'custom'",
        textAreaInput(
          "custom_query",
          "自定义 PubMed 检索式",
          placeholder = "例如：(melanoma) AND (CTC) AND (2020/01/01[Date - Publication] : 2024/12/31[Date - Publication])",
          rows = 5
        )
      ),
      fluidRow(
        column(6, numericInput("start_year", "起始年份", value = 2020, min = 1900, max = 2100)),
        column(6, numericInput("end_year", "结束年份", value = as.integer(format(Sys.Date(), "%Y")), min = 1900, max = 2100))
      ),
      selectInput(
        "article_type",
        "文献类型",
        choices = c("不限", "Journal Article", "Clinical Trial", "Randomized Controlled Trial", "Meta-Analysis", "Review", "Systematic Review"),
        selected = "不限"
      ),
      selectInput(
        "language",
        "语言",
        choices = c("不限", "英文", "中文", "其他"),
        selected = "不限"
      ),
      numericInput("retmax", "抓取条数", value = 20, min = 1, max = 10000),
      numericInput("timeout_sec", "超时秒数", value = 60, min = 10, max = 600),
      checkboxInput("include_abstract", "同时抓取摘要", value = TRUE),
      tags$hr(),
      checkboxInput("do_match", "匹配 JCR / IF", value = FALSE),
      conditionalPanel(
        condition = "input.do_match == true",
        fileInput("jcr_file", "上传 JCR 对照文件（TXT/TSV/XLSX，需含 ISO 和 JCR 列）", accept = c(".txt", ".tsv", ".xlsx", ".xls")),
        fileInput("if_file", "上传 IF 文件（XLSX，需含 JCRAbbreviation 和 JIF 列）", accept = c(".xlsx", ".xls"))
      ),
      tags$hr(),
      textInput("output_name", "导出文件名", value = "pubmed_results.xlsx"),
      passwordInput("api_key", "NCBI API Key（可选）", value = ""),
      actionButton("run_search", "开始抓取", class = "btn-primary"),
      br(), br(),
      downloadButton("download_result", "下载结果")
    ),
    mainPanel(
      h4("当前检索式"),
      verbatimTextOutput("query_text"),
      h4("运行状态"),
      verbatimTextOutput("status_text"),
      h4("结果预览"),
      tableOutput("result_preview")
    )
  )
)

server <- function(input, output, session) {
  rv <- reactiveValues(
    result_df = NULL,
    query = NULL,
    status = "等待开始",
    temp_file = NULL
  )

  observeEvent(input$run_search, {
    rv$status <- "正在构建检索式..."
    output$status_text <- renderText(rv$status)

    tryCatch({
      if (nzchar(trimws(input$api_key))) {
        Sys.setenv(ENTREZ_KEY = trimws(input$api_key))
      }

      query <- build_pubmed_query(
        keywords = parse_keywords(input$keywords),
        start_year = input$start_year,
        end_year = input$end_year,
        join_mode = input$join_mode,
        article_type = input$article_type,
        language = input$language,
        custom_query = input$custom_query
      )
      rv$query <- query

      rv$status <- "正在检索 PubMed 并抓取结果..."
      output$status_text <- renderText(rv$status)
      df <- fetch_pubmed_results(
        query = query,
        retmax = input$retmax,
        timeout_sec = input$timeout_sec,
        include_abstract = isTRUE(input$include_abstract)
      )

      if (isTRUE(input$do_match)) {
        if (is.null(input$jcr_file) || is.null(input$if_file)) {
          stop("已勾选 JCR / IF 匹配，但尚未上传完整文件。")
        }
        rv$status <- "正在匹配 JCR / IF..."
        output$status_text <- renderText(rv$status)
        df <- match_jcr_if(df, input$jcr_file, input$if_file)
      }

      rv$result_df <- df

      out_name <- trimws(input$output_name)
      if (!nzchar(out_name)) out_name <- "pubmed_results.xlsx"
      if (!grepl("\\.xlsx$", out_name, ignore.case = TRUE)) {
        out_name <- paste0(out_name, ".xlsx")
      }

      out_file <- file.path(tempdir(), out_name)
      save_results_excel(df, out_file, query)
      rv$temp_file <- out_file

      rv$status <- paste0("抓取完成。共获取 ", nrow(df), " 条记录。可直接下载结果。")
    }, error = function(e) {
      rv$status <- paste0("运行失败：", conditionMessage(e))
      rv$result_df <- NULL
      rv$temp_file <- NULL
    })
  })

  output$query_text <- renderText({
    if (is.null(rv$query)) "尚未生成检索式" else rv$query
  })

  output$status_text <- renderText({
    rv$status
  })

  output$result_preview <- renderTable({
    req(rv$result_df)
    head(rv$result_df, 20)
  }, striped = TRUE, bordered = TRUE, spacing = "s")

  output$download_result <- downloadHandler(
    filename = function() {
      out_name <- trimws(input$output_name)
      if (!nzchar(out_name)) out_name <- "pubmed_results.xlsx"
      if (!grepl("\\.xlsx$", out_name, ignore.case = TRUE)) {
        out_name <- paste0(out_name, ".xlsx")
      }
      out_name
    },
    content = function(file) {
      req(rv$temp_file)
      file.copy(rv$temp_file, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
