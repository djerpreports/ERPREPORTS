ALTER PROCEDURE [dbo].[spPage_GetPageList](
--DECLARE
    @GroupLabel			VARCHAR(100) --= ''
  , @PageName			VARCHAR(100) --= ''
  , @PageLabel			VARCHAR(100) --= ''
  , @URL				VARCHAR(100) --= ''
  , @HasSub				VARCHAR(100) --= ''
  , @ParentMenu			VARCHAR(100) --= ''
  , @ParentOrder		VARCHAR(100) --= ''
  , @Order				VARCHAR(100) --= ''

  , @StartPage			INT			--= 1
  , @RowCount			INT			--= 10
  , @SearchValue		VARCHAR(50)	--= ''
  , @SortColumnName		VARCHAR(50)	--= 'GroupLabel '
  , @SortDirection		VARCHAR(10)	--= 'ASC '
  , @TotalRecords		INT			OUTPUT
  , @FilteredRecords	INT			OUTPUT
)AS
BEGIN
  DECLARE @ListQuery NVARCHAR(max)
  DECLARE @ListFilterQuery NVARCHAR(max)
    SET @ListQuery = 'SELECT TOP('+ CAST(@RowCount AS NVARCHAR(5)) + ') * FROM ('
				+ ' SELECT *, ROW_NUMBER() OVER (ORDER BY ' + @SortColumnName + ' ' + @SortDirection + ') AS row_num '
                + ' FROM mPage'
                              + ' WHERE IsDeleted = 0 '
                              + ' AND (GroupLabel LIKE '''+'%'+ @GroupLabel +'%'''
                              + ' AND PageName LIKE '''+'%'+ @PageName +'%'''
                              + ' AND PageLabel LIKE '''+'%'+ @PageLabel +'%'''
                              + ' AND URL LIKE '''+'%'+ @URL +'%'''
                              + ' AND HasSub LIKE '''+'%'+ @HasSub +'%'''
                              + ' AND ParentMenu LIKE '''+'%'+ @ParentMenu +'%'''
                              + ' AND ParentOrder LIKE '''+'%'+ @ParentOrder +'%'''
                              + ' AND [Order] LIKE '''+'%'+ @Order + '%'')'
                              + ' AND (GroupLabel LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR PageName LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR PageLabel LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR URL LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR HasSub LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR ParentMenu LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR ParentOrder LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR [Order] LIKE '''+'%'+ @SearchValue + '%'')'                              
							  --+ ' ORDER BY '+ @SortColumnName+ ' '+ @SortDirection
                              + ') AS t WHERE row_num > ' + CAST( ( (@StartPage-1) * @RowCount) AS NVARCHAR(5))
    --SELECT @ListQuery
    
    EXEC sp_executesql @ListQuery
    
    SELECT @FilteredRecords = @@ROWCOUNT  
/*
  SET @ListFilterQuery = 'SET @FilteredRecords = ( SELECT COUNT(*) '
                + ' FROM mPage'
                              + ' WHERE IsDeleted = 0 '
                              + ' AND (GroupLabel LIKE '''+'%'+ @GroupLabel +'%'''
                              + ' AND PageName LIKE '''+'%'+ @PageName +'%'''
                              + ' AND PageLabel LIKE '''+'%'+ @PageLabel +'%'''
                              + ' AND URL LIKE '''+'%'+ @URL +'%'''
                              + ' AND HasSub LIKE '''+'%'+ @HasSub +'%'''
                              + ' AND ParentMenu LIKE '''+'%'+ @ParentMenu +'%'''
                              + ' AND ParentOrder LIKE '''+'%'+ @ParentOrder +'%'''
                              + ' AND Order LIKE '''+'%'+ @Order + '%'')'
                              + ' AND (GroupLabel LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR PageName LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR PageLabel LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR URL LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR HasSub LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR ParentMenu LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR ParentOrder LIKE '''+'%'+ @SearchValue +'%'''
                              + ' OR Order LIKE '''+'%'+ @SearchValue+'%'')' 
                              
              
    
    --PREPARE filterQueryStmt FROM @ListFilterQuery;
  --EXECUTE filterQueryStmt;
  EXECUTE @ListFilterQuery
  --DEALLOCATE PREPARE filterQueryStmt;  
    */
    SET @TotalRecords = (SELECT COUNT(*)
						 FROM mPage
                         WHERE IsDeleted = 0);
  
  --SELECT @TotalRecords		, @FilteredRecords
  
END
