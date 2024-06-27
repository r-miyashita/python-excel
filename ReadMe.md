## 変数
|  |  |  |  |
| - | - | - | - |
|  | inputNum | int | ユーザ入力 |
|  | params | dict | 定義情報( from: json ) |
|  | err_reasons | dict | エラー内容 |
|  | root |  |  |
|  | input_dir |  |  |
|  | input_files |  |  |
|  | out_put_dir |  |  |
|  | output_file |  |  |
|  | table |  |  |
| key_val_dict | updt_clmns | dict |  |
|  | ws_list | list |  |
| key_idxs | updt_clmns_idx | list | 更新カラムのインデックス |
| keys | updt_clmn_names | list | 更新カラム名 |
| vals_per_ws | updt_src | list | 更新情報 |
| key_addrs | updt_src_cells | list | 更新カラム名のセル情報 |
| val_addrs | updt_src_cell_vals | list | 更新値のセル情報 |
| append_col_no | append_column_no | int | 追加列の番号 |
| set_key_addrs | colname_cells | list | sql_set用カラム名のセル番号 |
| keys_info | colnames | list | カラム名情報（マッチングで使う） |
| set_key_addr | col_pos | str | カラム名参照用セル番地 |
| set_val_addr | val_pos | str | 値参照用セル番地 |
| cond_key_addr | condition_col | str | 条件カラム名のセル番地 |
| cond_val_addr | condition_val | str | 条件カラム値のセル番地 |