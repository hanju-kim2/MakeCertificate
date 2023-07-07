module XMLDefine

open FSharp.Data

type FileConfig = XmlProvider<"""
<FilePath><Reception>Path</Reception><Data>Path</Data><TestPhoto>Path</TestPhoto><ProductPhoto>Path</ProductPhoto><Report>Path</Report>
<ExcelSheetMatch><Excel><Name>Name</Name><Sheet>Sheet</Sheet></Excel><Excel><Name>Name</Name><Sheet>Sheet</Sheet></Excel></ExcelSheetMatch></FilePath>""">


type CellConfig = XmlProvider<"""
<Cells>
<BasicCells>
  <BasicCell>
    <FromFile>File</FromFile>
    <FromCell>Cell</FromCell>
    <ToCell>Cell</ToCell>
  </BasicCell>
  <BasicCell>
    <FromFile>Cell</FromFile>
    <FromCell>Cell</FromCell>
    <ToCell>Cell</ToCell>
  </BasicCell>
</BasicCells>
 
<AlternativeCells>
  <AlternativeCell>
    <FromFile>File</FromFile>
    <FromCell1>Cell</FromCell1>
    <FromCell2>Cell</FromCell2>
    <ToCell>Cell</ToCell>
  </AlternativeCell>
  <AlternativeCell>
    <FromFile>File</FromFile>
    <FromCell1>Cell</FromCell1>
    <FromCell2>Cell</FromCell2>
    <ToCell>Cell</ToCell>
  </AlternativeCell>
</AlternativeCells>

<CombineCells>
  <CombineCell>
    <FromFile>File</FromFile>
    <FromCell1>Cell</FromCell1>
    <FromCell2>Cell</FromCell2>
    <Format>format</Format>
    <ToCell>Cell</ToCell>
  </CombineCell>
  <CombineCell>
    <FromFile>File</FromFile>
    <FromCell1>Cell</FromCell1>
    <FromCell2>Cell</FromCell2>
    <Format>format</Format>
    <ToCell>Cell</ToCell>
  </CombineCell>
</CombineCells>


<ConditionCells>
    <ConditionCell>
      <FromFile>측정지시서</FromFile>
      <FromName>Check Box 31</FromName>
      <FromCells>
        <FromCell>G2</FromCell>
        <FromCell>G3</FromCell>
        <FromCell>G4</FromCell>
      </FromCells>
      <ToCells>
        <ToCell>F2539</ToCell>
        <ToCell>F2540</ToCell>
        <ToCell>F2541</ToCell>
      </ToCells>
    </ConditionCell>
    <ConditionCell>
        <FromFile>측정지시서</FromFile>
        <FromName>Check Box 31</FromName>
        <FromCells>
          <FromCell>G2</FromCell>
          <FromCell>G3</FromCell>
          <FromCell>G4</FromCell>
        </FromCells>
        <ToCells>
          <ToCell>F2539</ToCell>
          <ToCell>F2540</ToCell>
          <ToCell>F2541</ToCell>
        </ToCells>
    </ConditionCell>
</ConditionCells>

<NewPageLists>
    <NewPageList>
        <FromFile>엑셀.xlsx</FromFile>
        <FromCell>A2</FromCell>
        <Page>24</Page>
    </NewPageList>
    <NewPageList>
        <FromFile>엑셀.xlsx</FromFile>
        <FromCell>A2</FromCell>
        <Page>24</Page>
    </NewPageList>
</NewPageLists>

<ConditionOrderCells>
    <ConditionOrderCell>
       <FromFile>엑셀.xlsx</FromFile>
       <InputCols>A,D,K,AA,AD,AI</InputCols>
       <OutputCols>A,D,K,AA,AD,AI</OutputCols>
       <InputCheckBoxes>
         <CheckBox>
            <Name>Check box 34</Name>
            <Line>32</Line>
         </CheckBox>
         <CheckBox>
            <Name>Check box 34</Name>
            <Line>32</Line>
         </CheckBox>
       </InputCheckBoxes>
       <OutputCheckBoxes>
         <CheckBox>
            <Name>Check box 34</Name>
            <Line>32</Line>
         </CheckBox>
         <CheckBox>
            <Name>Check box 34</Name>
            <Line>32</Line>
         </CheckBox>
       </OutputCheckBoxes>
   </ConditionOrderCell>
   <ConditionOrderCell>
   <FromFile>엑셀.xlsx</FromFile>
   <InputCols>A,D,K,AA,AD,AI</InputCols>
   <OutputCols>A,D,K,AA,AD,AI</OutputCols>
   <InputCheckBoxes>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
   </InputCheckBoxes>
   <OutputCheckBoxes>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
   </OutputCheckBoxes>
   </ConditionOrderCell>
</ConditionOrderCells>

<ConditionOrderCheckCells>
<ConditionOrderCheckCell>
   <FromFile>엑셀.xlsx</FromFile>
   <FromName>Check Box 487</FromName>
   <InputCols>A,D,K,AA,AD,AI</InputCols>
   <OutputCols>A,D,K,AA,AD,AI</OutputCols>
   <InputCheckBoxes>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
   </InputCheckBoxes>
   <OutputCheckBoxes>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
     <CheckBox>
        <Name>Check box 34</Name>
        <Line>32</Line>
     </CheckBox>
   </OutputCheckBoxes>
   </ConditionOrderCheckCell>

   <ConditionOrderCheckCell>
   <FromFile>엑셀.xlsx</FromFile>
   <FromName>Check Box 487</FromName>
   <InputCols>A,D,K,AA,AD,AI</InputCols>
   <OutputCols>A,D,K,AA,AD,AI</OutputCols>
   <InputCheckBoxes>
 <CheckBox>
    <Name>Check box 34</Name>
    <Line>32</Line>
 </CheckBox>
 <CheckBox>
    <Name>Check box 34</Name>
    <Line>32</Line>
 </CheckBox>
   </InputCheckBoxes>
   <OutputCheckBoxes>
 <CheckBox>
    <Name>Check box 34</Name>
    <Line>32</Line>
 </CheckBox>
 <CheckBox>
    <Name>Check box 34</Name>
    <Line>32</Line>
 </CheckBox>
   </OutputCheckBoxes>
   </ConditionOrderCheckCell>
</ConditionOrderCheckCells>

<AllCompositions>
<AllComposition>
<FromFile>측정지시서</FromFile>
     <InputCols>B,D,E,G</InputCols>
     <OutputCols>B,J,R,V</OutputCols>
<InputCheckBoxes>
  <CheckBox>
     <Name>Check Box 234</Name>
     <Line>90</Line>
  </CheckBox>
  <CheckBox>
     <Name>Check Box 235</Name>
     <Line>91</Line>
  </CheckBox>
</InputCheckBoxes>
<OutputLines>
     <Line>782</Line>
     <Line>784</Line>
</OutputLines>
   </AllComposition>
   <AllComposition>
      <FromFile>측정지시서</FromFile>
     <InputCols>B,D,E,G</InputCols>
     <OutputCols>B,J,R,V</OutputCols>
      <InputCheckBoxes>
        <CheckBox>
           <Name>Check Box 234</Name>
           <Line>90</Line>
        </CheckBox>
        <CheckBox>
           <Name>Check Box 235</Name>
           <Line>91</Line>
        </CheckBox>
      </InputCheckBoxes>
      <OutputLines>
           <Line>782</Line>
           <Line>784</Line>
      </OutputLines>
   </AllComposition>

</AllCompositions>


<OrderCells>
  <OrderCell>
   <FromFile>측정지시서</FromFile>
   <InputCols>A,D,E,F,G,H</InputCols>
   <OutputCols>B,I,N,T,Z,AD</OutputCols>
   <InputLines>
      <Line>546</Line>
      <Line>548</Line>
      <Line>550</Line>
      <Line>552</Line>
      <Line>554</Line>
      <Line>556</Line>
   </InputLines>
   <OutputLines>
      <Line>546</Line>
      <Line>548</Line>
      <Line>550</Line>
   </OutputLines>
   </OrderCell>

   <OrderCell>
    <FromFile>측정지시서</FromFile>
    <InputCols>A,D,E,F,G,H</InputCols>
    <OutputCols>B,I,N,T,Z,AD</OutputCols>
    <InputLines>
       <Line>546</Line>
       <Line>548</Line>
       <Line>550</Line>
       <Line>552</Line>
       <Line>554</Line>
       <Line>556</Line>
    </InputLines>
    <OutputLines>
       <Line>546</Line>
       <Line>548</Line>
       <Line>550</Line>
    </OutputLines>
    </OrderCell>
</OrderCells>


<DeletePages>
<DeletePage>
  <Page>21</Page>
  <Flag>C12</Flag>
</DeletePage>
<DeletePage>
    <Page>21</Page>
    <Flag>C12</Flag>
</DeletePage>
</DeletePages>

</Cells>

"""
>


type ControlConfig = XmlProvider<"""
<Controls>
<CheckSingles>
  <CheckSingle>
    <FromFile>EMC-001(Ver 5)_test data sheet_조명기기용_전도 고색 Ver_191002 하모닉 교정.xlsx</FromFile>
    <FromName>G26</FromName>
    <ToName>S225</ToName>
  </CheckSingle>
  <CheckSingle>
    <FromFile>EMC-001(Ver 5)_test data sheet_조명기기용_전도 고색 Ver_191002 하모닉 교정.xlsx</FromFile>
    <FromName>G26</FromName>
    <ToName>AA225</ToName>
  </CheckSingle>
  <CheckSingle>
    <FromFile>EMC-001(Ver 5)_test data sheet_조명기기용_전도 고색 Ver_191002 하모닉 교정.xlsx</FromFile>
    <FromName>G26</FromName>
    <ToName>S227</ToName>
  </CheckSingle>
  <CheckSingle>
    <FromFile>EMC-001(Ver 5)_test data sheet_조명기기용_전도 고색 Ver_191002 하모닉 교정.xlsx</FromFile>
    <FromName>H26</FromName>
    <ToName>A227</ToName>
  </CheckSingle>
  <CheckSingle>
    <FromFile>EMC-001(Ver 5)_test data sheet_조명기기용_전도 고색 Ver_191002 하모닉 교정.xlsx</FromFile>
    <FromName>H26</FromName>
    <ToName>S229</ToName>
  </CheckSingle>
</CheckSingles>

<CheckGroups>
  <CheckGroup>
    <FromFile>시험평가신청서.xlsx</FromFile>
    <FromNames>
      <FromName>AG6</FromName>
      <FromName>AG6</FromName>
    </FromNames>
    <ToName>J10</ToName>
  </CheckGroup>
  <CheckGroup>
  <FromFile>시험평가신청서.xlsx</FromFile>
  <FromNames>
    <FromName>AG6</FromName>
    <FromName>AG6</FromName>
  </FromNames>
  <ToName>J10</ToName>
  </CheckGroup>
</CheckGroups>

<InverseCheckSingles>
    <InverseCheckSingle>
        <FromFile>측정지시서</FromFile>
        <FromName>Check Box 3</FromName>
        <ToName>확인란 160</ToName>
    </InverseCheckSingle>
    <InverseCheckSingle>
        <FromFile>측정지시서</FromFile>
        <FromName>Check Box 3</FromName>
        <ToName>확인란 160</ToName>
    </InverseCheckSingle>
</InverseCheckSingles>

<ValueExists>
<ValueExist>
  <FromFile>인풋.xlsx</FromFile>
  <FromCell>G26</FromCell>
  <ToName>확인란 160</ToName>
</ValueExist>
<ValueExist>
    <FromFile>인풋.xlsx</FromFile>
    <FromCell>G26</FromCell>
    <ToName>확인란 160</ToName>
</ValueExist>
</ValueExists>

<InverseCheckGroups>
  <InverseCheckGroup>
    <FromFile>측정지시서</FromFile>
    <FromNameInverse>Check Box 485</FromNameInverse>
    <FromName>Check Box 29</FromName>
    <ToName>Check Box 141</ToName>
  </InverseCheckGroup>
  <InverseCheckGroup>
    <FromFile>측정지시서</FromFile>
    <FromNameInverse>Check Box 487</FromNameInverse>
    <FromName>Check Box 29</FromName>
    <ToName>Check Box 190</ToName>
  </InverseCheckGroup>
</InverseCheckGroups>

</Controls>
""">


type PhotoConfig = XmlProvider<"""
<Photos>
<BasicPhotos>
  <BasicPhoto>
  <Photo>asdf.jpg</Photo>
  <ToCellStart>J10</ToCellStart>
  <ToCellEnd>N15</ToCellEnd>
  <Flag>C12</Flag>
  </BasicPhoto>
  <BasicPhoto>
  <Photo>asdf.jpg</Photo>
  <ToCellStart>J10</ToCellStart>
  <ToCellEnd>N15</ToCellEnd>
  <Flag>C12</Flag>
  </BasicPhoto>
</BasicPhotos>
<ExcelPhotos>
  <ExcelPhoto>
   <Excel>시험평가신청서.xlsx</Excel>
   <Photo>그림1</Photo>
   <ToCellStart>B3258</ToCellStart>
   <ToCellEnd>K3300</ToCellEnd>
   <Flag>B3258</Flag>
  </ExcelPhoto>
  <ExcelPhoto>
   <Excel>시험평가신청서.xlsx</Excel>
   <Photo>그림1</Photo>
   <ToCellStart>B3258</ToCellStart>
   <ToCellEnd>K3300</ToCellEnd>
   <Flag>B3258</Flag>
   </ExcelPhoto>
</ExcelPhotos>
<NewPagePhotos>
 <NewPagePhoto>
  <Tag>CE1</Tag>
  <Page>14</Page>
  <PhotoPerPage>2</PhotoPerPage>
 </NewPagePhoto>
 <NewPagePhoto>
  <Tag>CE1</Tag>
  <Page>14</Page>
  <PhotoPerPage>2</PhotoPerPage>
 </NewPagePhoto>
</NewPagePhotos>
<ConditionPhotos>
  <ConditionPhoto>
    <FromFile>시험평가신청서.xlsx</FromFile>
    <FromName>Check Box 484</FromName>
    <Photo>CE1.JPG</Photo>
    <ToCellStart>B3238</ToCellStart>
    <ToCellEnd>K3300</ToCellEnd>
    <Flag>B3238</Flag>
  </ConditionPhoto>
  <ConditionPhoto>
    <FromFile>시험평가신청서.xlsx</FromFile>
    <FromName>Check Box 484</FromName>
    <Photo>CE1.JPG</Photo>
    <ToCellStart>B3238</ToCellStart>
    <ToCellEnd>K3300</ToCellEnd>
    <Flag>B3238</Flag>
  </ConditionPhoto>
</ConditionPhotos>
</Photos>
""">


type ErrorConfig = XmlProvider<"""
<Errors>
    <ErrorCells>
        <ErrorCell>
            <FromFile>시험평가신청서.xlsx</FromFile>
            <Date>AG6</Date>
            <Pivot>AG6</Pivot>
        </ErrorCell>
        <ErrorCell>
            <FromFile>시험평가신청서.xlsx</FromFile>
            <Date>AG6</Date>
            <Pivot>NOW</Pivot>
        </ErrorCell>
    </ErrorCells>

</Errors>
""">
