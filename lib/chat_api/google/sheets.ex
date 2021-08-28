defmodule ChatApi.Google.Sheets do
  def get_spreadsheet_info(refresh_token, id) do
    scope = "https://sheets.googleapis.com/v4/spreadsheets/#{id}"
    client = ChatApi.Google.Auth.get_token!(refresh_token: refresh_token)

    case OAuth2.Client.get(client, scope) do
      {:ok, %{body: result}} -> {:ok, result}
      error -> error
    end
  end

  def get_spreadsheet_values(refresh_token, id) do
    with {:ok, %{"sheets" => _} = result} <- get_spreadsheet_info(refresh_token, id),
         {:ok, [default_sheet_name | _]} <- extract_sheet_names(result) do
      range = "#{default_sheet_name}!A:Z"
      scope = "https://sheets.googleapis.com/v4/spreadsheets/#{id}/values/#{range}"
      client = ChatApi.Google.Auth.get_token!(refresh_token: refresh_token)

      case OAuth2.Client.get(client, scope) do
        {:ok, %{body: result}} -> {:ok, result}
        error -> error
      end
    end
  end

  def get_spreadsheet_values(refresh_token, id, range) do
    scope = "https://sheets.googleapis.com/v4/spreadsheets/#{id}/values/#{range}"
    client = ChatApi.Google.Auth.get_token!(refresh_token: refresh_token)

    case OAuth2.Client.get(client, scope) do
      {:ok, %{body: result}} -> {:ok, result}
      error -> error
    end
  end

  def get_spreadsheet_by_id!(refresh_token, id, range \\ "Sheet1!A:Z") do
    scope = "https://sheets.googleapis.com/v4/spreadsheets/#{id}/values/#{range}"
    client = ChatApi.Google.Auth.get_token!(refresh_token: refresh_token)
    %{body: result} = OAuth2.Client.get!(client, scope)

    result
  end

  def append_to_spreadsheet!(refresh_token, id, data \\ [], range \\ "Sheet1!A:Z") do
    qs = URI.encode_query(%{valueInputOption: "USER_ENTERED", includeValuesInResponse: true})
    scope = "https://sheets.googleapis.com/v4/spreadsheets/#{id}/values/#{range}:append?#{qs}"
    client = ChatApi.Google.Auth.get_token!(refresh_token: refresh_token)

    payload = %{
      "majorDimension" => "ROWS",
      "range" => range,
      "values" =>
        if Enum.all?(data, &is_list/1) do
          data
        else
          # If we're only inserting one row, make sure it's formatted properly
          [data]
        end
    }

    %{body: result} = OAuth2.Client.post!(client, scope, payload)

    result
  end

  def extract_sheet_names(%{"sheets" => sheets}) when is_list(sheets) do
    {:ok,
     Enum.map(sheets, fn sheet ->
       get_in(sheet, ["properties", "title"])
     end)}
  end

  def extract_sheet_names(_), do: {:error, "Unable to find sheets for spreadsheet!"}

  def format_as_json(%{"values" => values}) when is_list(values) do
    [headers | rows] = Enum.reject(values, &Enum.empty?/1)

    keys =
      Enum.map(headers, fn header ->
        header |> String.split(" ") |> Enum.join("_") |> String.downcase()
      end)

    Enum.map(rows, fn items ->
      items
      |> Enum.with_index()
      |> Enum.reduce(%{}, fn {value, index}, acc ->
        case Enum.at(keys, index) do
          nil -> acc
          key -> Map.merge(acc, %{key => value})
        end
      end)
    end)
  end

  def format_as_json(_response), do: []
end
